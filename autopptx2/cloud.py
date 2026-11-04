"""Đăng nhập + đồng bộ Excel / Avatar / ảnh nền với máy chủ kho của công ty.

Máy chủ: server/autopptx_server.py chạy trên VPS (thay Supabase). App chỉ giữ
phiên đăng nhập (token, thu hồi được), không lưu mật khẩu.

Không import Tkinter — gọi từ worker thread. App chỉ nhận dict kết quả
rồi cập nhật UI trên luồng chính.
"""
import os
import re
import ssl
import json
import hashlib
import datetime
import shutil
import sys
import threading
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

from .constants import (CONFIG_DIR, CONFIG_PATH, V1_CREDS_PATH, KEYRING_SERVICE,
                        SERVER_URL, IMG_EXTS)

KEYRING_IMPORT_ERROR = ''
try:
    import keyring
except Exception as _e:
    keyring = None
    KEYRING_IMPORT_ERROR = '{}: {}'.format(type(_e).__name__, _e)

try:
    import certifi
except Exception:
    certifi = None

# Số luồng tải song song (máy chủ nhẹ, mỗi ảnh ~0,3–1 MB).
SYNC_WORKERS = 4
TIMEOUT = 30


def runtime_info():
    """'bản cài · macOS 12.7 · x86_64 · Python 3.11.9' — để ảnh chụp lỗi tự đủ thông tin."""
    import platform
    where = 'bản cài' if getattr(sys, 'frozen', False) else 'chạy từ source'
    if sys.platform == 'darwin':
        osv = 'macOS ' + (platform.mac_ver()[0] or '?')
    elif sys.platform == 'win32':
        osv = 'Windows ' + (platform.release() or '?')
    else:
        osv = platform.system() or sys.platform
    return '{} · {} · {} · Python {}'.format(
        where, osv, platform.machine() or '?', platform.python_version())


def _client_label():
    """Nhãn máy cho User-Agent / danh sách phiên — header HTTP chỉ nhận ASCII."""
    try:
        from . import __version__ as v
    except Exception:
        v = '?'
    txt = unicodedata.normalize('NFKD', f'AutoPPTX {v} · {runtime_info()}'.replace('·', '-'))
    return txt.replace('đ', 'd').encode('ascii', 'ignore').decode('ascii')


_KEYRING_TIMEOUT = 2.5 if (sys.platform == 'darwin' or getattr(sys, 'frozen', False)) else 8.0


def _keyring_call(fn, *args):
    """Keychain macOS / app đóng gói có thể treo — giới hạn thời gian."""
    if not keyring:
        return None
    box = {'v': None}

    def work():
        try:
            box['v'] = fn(*args)
        except Exception:
            box['v'] = None

    t = threading.Thread(target=work, daemon=True)
    t.start()
    t.join(_KEYRING_TIMEOUT)
    if t.is_alive():
        print('keyring: timeout — bỏ qua Keychain, dùng file config nếu có')
        return None
    return box['v']


def _keyring_get_password(service, username):
    if not keyring or not username:
        return None
    return _keyring_call(keyring.get_password, service, username)


_AVATAR_STRIP = re.compile(r'(_Ava|AA|A)$', re.IGNORECASE)


def tool_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _writable_dir(*candidates):
    for d in candidates:
        try:
            os.makedirs(d, exist_ok=True)
            t = os.path.join(d, '.wtest')
            with open(t, 'w') as f:
                f.write('')
            os.remove(t)
            return d
        except Exception:
            continue
    d = os.path.join(CONFIG_DIR, 'cache')
    os.makedirs(d, exist_ok=True)
    return d


def excel_cache_dir():
    return _writable_dir(os.path.join(CONFIG_DIR, 'excel_cache'))


def excel_cache_path(user_id):
    if not user_id:
        return ''
    return os.path.join(excel_cache_dir(), f'{user_id}.xlsx')


def _utc_now_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def _ts_equal(a, b):
    """So sánh updated_at — bỏ qua khác Z vs +00:00 / thiếu múi giờ."""
    if not a and not b:
        return True
    if not a or not b:
        return False

    def norm(s):
        s = str(s).strip().replace('Z', '+00:00')
        if 'T' in s and '+' not in s[10:] and not s.endswith('-00:00'):
            s = s + '+00:00'
        return s

    return norm(a) == norm(b)


def avatar_cache_dir():
    return _writable_dir(os.path.join(tool_dir(), 'Data avatar'),
                         os.path.join(CONFIG_DIR, 'avatar_cache'))


def bg_cache_dir():
    return _writable_dir(os.path.join(tool_dir(), 'Data background'),
                         os.path.join(CONFIG_DIR, 'bg_cache'))


def _read_json(path, default=None):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default if default is not None else {}


def _write_json(path, data):
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


_SHA_RE = re.compile(r'^[0-9a-f]{64}$')


def _is_sha(v):
    return bool(v and _SHA_RE.match(str(v)))


def _nfc(s):
    return unicodedata.normalize('NFC', str(s or ''))


def _local_path(base, rel):
    return os.path.join(base, *rel.split('/'))


def _quote_path(rel):
    return '/'.join(urllib.parse.quote(p, safe='') for p in rel.split('/'))


# ── Ghi nhớ phiên đăng nhập (token, KHÔNG phải mật khẩu) ──
def _token_user(email):
    return 'session:' + (email or '').strip().lower()


def load_saved_session():
    """{'email', 'token'} nếu máy này đã chọn «Giữ đăng nhập». token có thể rỗng."""
    d = _read_json(CONFIG_PATH)
    email = d.get('email') or ''
    if not email:
        return None
    token = d.get('session_token') or ''
    if not token:
        token = _keyring_get_password(KEYRING_SERVICE, _token_user(email)) or ''
    return {'email': email, 'token': token}


def save_session(email, token):
    stored = False
    if keyring and token:
        _keyring_call(keyring.set_password, KEYRING_SERVICE, _token_user(email), token)
        stored = _keyring_get_password(KEYRING_SERVICE, _token_user(email)) == token
    d = _read_json(CONFIG_PATH)
    d['email'] = email
    d['remember'] = True
    if stored:
        d.pop('session_token', None)
    else:
        d['session_token'] = token
    d.pop('pw', None)
    _write_json(CONFIG_PATH, d)
    forget_legacy_password(email)


def clear_session(keep_email=True):
    d = _read_json(CONFIG_PATH)
    email = d.get('email', '')
    if keyring and email:
        _keyring_call(keyring.delete_password, KEYRING_SERVICE, _token_user(email))
    d.pop('session_token', None)
    d.pop('pw', None)
    d.pop('remember', None)
    if not keep_email:
        d.pop('email', None)
    _write_json(CONFIG_PATH, d)


def forget_legacy_password(email=None):
    """Bản cũ (Supabase) nhớ mật khẩu thật trong keyring / file base64 — xoá đi."""
    for path in (CONFIG_PATH, V1_CREDS_PATH):
        d = _read_json(path)
        if not d:
            continue
        em = d.get('email') or email
        changed = False
        if 'pw' in d:
            d.pop('pw', None)
            changed = True
        if changed:
            try:
                _write_json(path, d)
            except Exception:
                pass
        if keyring and em:
            _keyring_call(keyring.delete_password, KEYRING_SERVICE, em)


class CloudError(Exception):
    def __init__(self, msg, status=0):
        super().__init__(msg)
        self.msg = msg
        self.status = status


class CloudClient:
    def __init__(self, url=None):
        self.url = (url or SERVER_URL).strip().rstrip('/')
        self.token = None
        self._api_lock = threading.Lock()
        self.user_id = None
        self.email = None
        self.role = 'user'
        self.name = None
        self.must_change = False
        self._ssl = None

    # ── trạng thái ──
    def is_admin(self):
        return self.role in ('admin', 'super_admin')

    def logged_in(self):
        return bool(self.user_id and self.token)

    def package_ok(self):
        return True

    def admin_url(self):
        return self.url + '/admin/'

    def web_url(self):
        return self.url + '/'

    def save_preset(self, name, layout, opts, shared=False):
        """Lưu bố cục + tuỳ chọn hiện tại thành mẫu trên máy chủ (bản web dùng)."""
        from .pipeline import portable_preset
        try:
            d = self._request('PUT', '/api/presets', body={
                'name': name, 'shared': bool(shared), 'data': portable_preset(layout, opts)})
        except CloudError as e:
            return {'ok': False, 'msg': e.msg}
        return {'ok': True, 'msg': d.get('msg') or 'Đã lưu mẫu lên máy chủ.'}

    # ── HTTP ──
    def _ctx(self):
        if self._ssl is None:
            try:
                self._ssl = (ssl.create_default_context(cafile=certifi.where())
                             if certifi else ssl.create_default_context())
            except Exception:
                self._ssl = ssl.create_default_context()
        return self._ssl

    def _request(self, method, path, body=None, headers=None, timeout=TIMEOUT,
                 want='json', dest=None, auth=True):
        """Gọi API. want='json' trả dict; want='file' ghi vào dest (atomic)."""
        h = {'User-Agent': _client_label(), 'Accept': 'application/json'}
        if auth and self.token:
            h['Authorization'] = 'Bearer ' + self.token
        h.update(headers or {})
        data = body
        if isinstance(body, dict):
            data = json.dumps(body).encode('utf-8')
            h['Content-Type'] = 'application/json'
        req = urllib.request.Request(self.url + path, data=data, method=method, headers=h)
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=self._ctx()) as resp:
                if want == 'file':
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                    tmp = dest + '.part'
                    with open(tmp, 'wb') as f:
                        shutil.copyfileobj(resp, f, 1 << 16)
                    os.replace(tmp, dest)
                    return {'ok': True, 'path': dest}
                raw = resp.read()
                return json.loads(raw.decode('utf-8') or '{}')
        except urllib.error.HTTPError as e:
            try:
                msg = json.loads(e.read().decode('utf-8')).get('msg') or ''
            except Exception:
                msg = ''
            if e.code == 401 and auth and path not in ('/api/login',):
                self.user_id = None
                msg = msg or 'Phiên đăng nhập đã hết. Đăng nhập lại.'
            raise CloudError(msg or f'Máy chủ trả lỗi {e.code}.', e.code)
        except urllib.error.URLError as e:
            reason = getattr(e, 'reason', e)
            if isinstance(reason, ssl.SSLError):
                raise CloudError(f'Lỗi chứng chỉ HTTPS khi kết nối máy chủ: {reason}')
            host = urllib.parse.urlsplit(self.url).netloc or self.url
            raise CloudError(f'Không kết nối được máy chủ {host}. Kiểm tra mạng rồi thử lại.')
        except (TimeoutError, OSError) as e:
            raise CloudError(f'Mất kết nối tới máy chủ: {e}')

    # ── Auth ──
    def _set_user(self, u):
        self.user_id = str(u.get('id')) if u.get('id') is not None else None
        self.email = u.get('email') or self.email
        self.role = u.get('role') or 'user'
        self.name = u.get('name') or None
        self.must_change = bool(u.get('must_change'))

    def login(self, email, password):
        """Trả (ok, message)."""
        self.user_id = self.name = self.token = None
        self.role = 'user'
        self.email = email
        try:
            d = self._request('POST', '/api/login', auth=False, body={
                'email': email, 'password': password, 'client': _client_label()})
        except CloudError as e:
            return False, e.msg
        self.token = d.get('token')
        self._set_user(d.get('user') or {})
        if not self.token or not self.user_id:
            return False, 'Máy chủ không trả phiên đăng nhập.'
        return True, 'ok'

    def resume(self, email, token):
        """Vào lại bằng phiên đã nhớ. Trả (ok, message)."""
        if not token:
            return False, ''
        self.token = token
        self.email = email
        try:
            d = self._request('GET', '/api/me')
        except CloudError as e:
            self.token = None
            self.user_id = None
            return False, e.msg
        self._set_user(d.get('user') or {})
        return bool(self.user_id), 'ok'

    def change_password(self, old, new):
        try:
            d = self._request('POST', '/api/me/password', body={'old': old, 'new': new})
        except CloudError as e:
            return {'ok': False, 'msg': e.msg}
        self.must_change = False
        return {'ok': True, 'msg': d.get('msg') or 'Đã đổi mật khẩu.'}

    def logout(self):
        if self.token:
            try:
                self._request('POST', '/api/logout', timeout=8)
            except Exception:
                pass
        self.token = None
        self.user_id = self.email = self.name = None
        self.role = 'user'
        self.must_change = False

    # ── Excel ──
    def excel_meta_path(self):
        return os.path.join(excel_cache_dir(), 'sync_meta.json')

    def excel_status_text(self):
        if not self.user_id:
            return 'Chưa đăng nhập'
        meta = _read_json(self.excel_meta_path()).get(self.user_id) or {}
        if not meta:
            return 'Chưa đồng bộ lần nào'
        name = meta.get('file_name') or ''
        when = meta.get('synced_at_local', '')
        if name:
            return f'{name} · {when}' if when else name
        return f'Lần cuối: {when}' if when else 'Chưa đồng bộ lần nào'

    def excel_local_display_name(self, path):
        """Tên hiển thị — cache dùng file_name trong meta, không phải {uid}.xlsx."""
        if not path:
            return ''
        uid = self.user_id
        if uid:
            cache = excel_cache_path(uid)
            if (cache and os.path.normcase(os.path.normpath(path))
                    == os.path.normcase(os.path.normpath(cache))):
                meta = _read_json(self.excel_meta_path()).get(uid) or {}
                if meta.get('file_name'):
                    return meta['file_name']
        return os.path.basename(path)

    def _remember_excel(self, uid, info):
        meta = _read_json(self.excel_meta_path())
        meta[uid] = {
            'sha': info.get('sha'),
            'updated_at': info.get('updated_at'),
            'synced_at_local': datetime.datetime.now().strftime('%d/%m/%Y %H:%M'),
            'file_name': info.get('file_name') or '',
        }
        _write_json(self.excel_meta_path(), meta)

    def sync_excel(self, progress=None, force=False):
        """Tải FILE TỔNG áp cho tài khoản đang đăng nhập. Trả dict kết quả."""
        uid = self.user_id
        if not self.logged_in():
            return {'ok': False, 'msg': 'Cần đăng nhập.'}
        if progress:
            progress(0.15, 'Đang kiểm tra bản Excel mới…')
        try:
            info = self._request('GET', '/api/excel/meta').get('excel') or {}
        except CloudError as e:
            if e.status == 404:
                return {'ok': False, 'msg': 'Chưa có FILE TỔNG trên máy chủ cho tài khoản này.'}
            return {'ok': False, 'msg': f'Lỗi kết nối: {e.msg}'}
        cache_path = excel_cache_path(uid)
        local = _read_json(self.excel_meta_path()).get(uid, {})
        if (not force and local.get('sha') and local.get('sha') == info.get('sha')
                and os.path.exists(cache_path)):
            if progress:
                progress(1.0, 'Đã là bản mới nhất')
            return {'ok': True, 'path': cache_path, 'file_name': info.get('file_name'),
                    'fresh': False, 'msg': 'Excel đã là bản mới nhất.'}
        if progress:
            progress(0.6, 'Đang tải Excel…')
        try:
            self._request('GET', '/api/excel/file', want='file', dest=cache_path, timeout=120)
        except CloudError as e:
            return {'ok': False, 'msg': f'Lỗi tải Excel: {e.msg}'}
        self._remember_excel(uid, info)
        if progress:
            progress(1.0, 'Đã cập nhật Excel')
        return {'ok': True, 'path': cache_path, 'file_name': info.get('file_name'),
                'fresh': True, 'msg': 'Đã tải Excel từ máy chủ.'}

    def _put_excel(self, local_path, target, progress=None):
        def prog(frac, text):
            if progress:
                progress(frac, text)

        if not self.logged_in():
            return {'ok': False, 'msg': 'Cần đăng nhập.'}
        if not local_path or not os.path.exists(local_path):
            return {'ok': False, 'msg': 'Chưa chọn file Excel trên máy.'}
        file_name = os.path.basename(local_path)
        try:
            prog(0.12, 'Đang đọc file…')
            with open(local_path, 'rb') as f:
                data = f.read()
            prog(0.4, 'Đang tải lên máy chủ…')
            d = self._request(
                'PUT', '/api/excel?target=' + urllib.parse.quote(str(target)), body=data,
                headers={'X-File-Name': urllib.parse.quote(_nfc(file_name)),
                         'Content-Type': 'application/vnd.openxmlformats-officedocument'
                                         '.spreadsheetml.sheet'},
                timeout=300)
        except CloudError as e:
            return {'ok': False, 'msg': e.msg}
        except OSError as e:
            return {'ok': False, 'msg': f'Không đọc được file: {e}'}
        info = d.get('excel') or {}
        uid = self.user_id
        cache_path = excel_cache_path(uid)
        # Bản vừa đẩy áp cho chính mình (me / all) → cập nhật cache để khỏi tải lại.
        if target in ('me', 'all') and cache_path:
            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            shutil.copy2(local_path, cache_path)
            self._remember_excel(uid, info)
        prog(1.0, 'Đã đẩy xong')
        return {'ok': True, 'msg': d.get('msg') or 'Đã đẩy Excel lên máy chủ.',
                'path': cache_path or local_path, 'file_name': file_name}

    def upload_excel(self, local_path, target_user_id=None, progress=None):
        return self._put_excel(local_path, target_user_id or 'me', progress)

    def upload_excel_broadcast(self, local_path, profiles=None, progress=None):
        """Admin: một file chung cho mọi tài khoản (máy chủ bỏ các bản riêng)."""
        if not self.is_admin():
            return {'ok': False, 'msg': 'Chỉ quản trị mới đẩy cho mọi tài khoản.'}
        return self._put_excel(local_path, 'all', progress)

    # ── Avatar ──
    def avatar_meta_path(self):
        return os.path.join(avatar_cache_dir(), 'sync_meta.json')

    @staticmethod
    def _count_local_avatars(folder):
        n = 0
        if not folder or not os.path.isdir(folder):
            return 0
        for _r, _ds, files in os.walk(folder):
            n += sum(1 for f in files if f.lower().endswith(IMG_EXTS))
        return n

    def avatar_status_text(self):
        meta = _read_json(self.avatar_meta_path())
        n = meta.get('local_count')
        if n is None:
            n = self._count_local_avatars(avatar_cache_dir())
        last = meta.get('synced_at', '')
        return f'{n} ảnh trong kho' + (f' · {last}' if last else '')

    def _manifest(self, kind):
        return self._request('GET', f'/api/manifest/{kind}', timeout=60).get('files') or []

    @staticmethod
    def _migrate_supabase_layout(cache, files, files_meta):
        """Bản Supabase tải avatar vào <cache>/shared/<…>; máy chủ mới dùng <cache>/<…>.
        Chuyển chỗ thay vì tải lại hàng trăm MB."""
        old_root = os.path.join(cache, 'shared')
        if not os.path.isdir(old_root):
            return 0
        moved = 0
        for f in files:
            rel = f['path']
            src = _local_path(old_root, rel)
            dst = _local_path(cache, rel)
            if os.path.isfile(src) and not os.path.exists(dst):
                try:
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    os.replace(src, dst)
                    moved += 1
                except OSError:
                    pass
        for k in [k for k in files_meta if k.startswith('shared/')]:
            files_meta.pop(k, None)
        for root, _ds, _fs in sorted(os.walk(old_root), key=lambda t: -len(t[0])):
            try:
                os.rmdir(root)
            except OSError:
                pass
        return moved

    def sync_avatars(self, progress=None):
        if not self.logged_in():
            return {'ok': False, 'msg': 'Cần đăng nhập.'}
        if progress:
            progress(0, 1, 'Đang kiểm tra kho avatar…')
        try:
            files = self._manifest('avatars')
        except CloudError as e:
            return {'ok': False, 'msg': f'Lỗi kết nối kho avatar: {e.msg}'}
        cache = avatar_cache_dir()
        meta = _read_json(self.avatar_meta_path())
        files_meta = meta.get('files', {})
        moved = self._migrate_supabase_layout(cache, files, files_meta)
        to_dl, skipped, adopted = [], 0, 0
        for f in files:
            rel, sha = f['path'], f['sha']
            local_fp = _local_path(cache, rel)
            if os.path.exists(local_fp):
                old = files_meta.get(rel)
                if old == sha:
                    skipped += 1
                    continue
                # Meta bản Supabase lưu updated_at (không phải sha): coi như chưa biết.
                if not _is_sha(old):
                    # Ảnh đã có sẵn trên máy (kho cũ / chép tay): nhận luôn, không tải lại.
                    files_meta[rel] = sha
                    adopted += 1
                    continue
            to_dl.append((rel, sha, local_fp))
        need = len(to_dl)
        state = {'done': 0, 'ok': 0, 'err': 0}
        lock = threading.Lock()

        def one(item):
            rel, sha, local_fp = item
            try:
                self._request('GET', '/api/files/avatars/' + _quote_path(rel),
                              want='file', dest=local_fp, timeout=120)
                good = True
            except CloudError as e:
                print('Avatar download err:', rel, e.msg)
                good = False
            with lock:
                state['done'] += 1
                if good:
                    files_meta[rel] = sha
                    state['ok'] += 1
                else:
                    state['err'] += 1
                i = state['done']
                if i % 100 == 0:
                    meta['files'] = files_meta
                    _write_json(self.avatar_meta_path(), meta)
            if progress and (i % 8 == 0 or i == need):
                progress(i, need, f'Đang tải {state["ok"]}/{need}…')

        if need:
            with ThreadPoolExecutor(max_workers=SYNC_WORKERS) as ex:
                list(ex.map(one, to_dl))
        meta['files'] = files_meta
        meta['synced_at'] = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        meta['local_count'] = self._count_local_avatars(cache)
        _write_json(self.avatar_meta_path(), meta)
        msg = (f'Avatar: kho máy chủ {len(files)} · tải {state["ok"]} mới, '
               f'sẵn có {skipped + adopted}' + (f', {state["err"]} lỗi' if state['err'] else ''))
        if moved:
            msg += f' · chuyển {moved} ảnh từ kho cũ'
        if progress:
            progress(1, 1, msg)
        return {'ok': state['err'] == 0 or state['ok'] > 0, 'path': cache,
                'downloaded': state['ok'], 'skipped': skipped + adopted,
                'errors': state['err'], 'msg': msg}

    def upload_avatars(self, folder, progress=None):
        if not self.logged_in():
            return {'ok': False, 'msg': 'Cần đăng nhập.'}
        if not self.is_admin():
            return {'ok': False, 'msg': 'Chỉ quản trị mới được đẩy avatar.'}
        files = []
        for root, _ds, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(IMG_EXTS) and not fn.startswith('.'):
                    files.append(os.path.join(root, fn))
        if not files:
            return {'ok': False, 'msg': 'Thư mục không có ảnh avatar.'}
        try:
            have = {f['path']: f.get('src_sha') for f in self._manifest('avatars')}
        except CloudError as e:
            return {'ok': False, 'msg': e.msg}
        state = {'done': 0, 'up': 0, 'same': 0, 'err': 0, 'last': ''}
        lock = threading.Lock()
        total = len(files)

        def one(path):
            rel = _nfc(os.path.relpath(path, folder).replace('\\', '/'))
            try:
                if have.get(rel) and have[rel] == _sha256_file(path):
                    res = 'same'
                else:
                    with open(path, 'rb') as f:
                        data = f.read()
                    d = self._request('PUT', '/api/files/avatars/' + _quote_path(rel),
                                      body=data, timeout=180,
                                      headers={'Content-Type': 'application/octet-stream'})
                    res = 'same' if d.get('unchanged') else 'up'
            except (CloudError, OSError) as e:
                res = 'err'
                state['last'] = f'{rel}: {getattr(e, "msg", e)}'
            with lock:
                state['done'] += 1
                state[res] += 1
                i = state['done']
            if progress:
                progress(i, total, f'Đẩy {i}/{total}…')

        with ThreadPoolExecutor(max_workers=SYNC_WORKERS) as ex:
            list(ex.map(one, files))
        msg = (f'Đã đẩy {state["up"]} avatar mới/đổi, {state["same"]} không đổi'
               + (f', {state["err"]} lỗi (vd {state["last"]})' if state['err'] else ''))
        return {'ok': state['err'] == 0, 'msg': msg}

    # ── Ảnh nền ──
    def list_backgrounds(self, prefix=''):
        if not self.logged_in():
            return []
        try:
            files = self._manifest('backgrounds')
        except CloudError as e:
            print('BG list err:', e.msg)
            return []
        out = []
        for f in files:
            sp = f['path']
            if prefix and not sp.startswith(prefix.rstrip('/') + '/'):
                continue
            # updated_at dùng làm «phiên bản» để cache: lấy luôn sha cho chắc.
            out.append({'storage_path': sp, 'file_name': sp.rsplit('/', 1)[-1],
                        'updated_at': f.get('sha') or f.get('updated_at') or ''})
        return out

    def bg_thumb_path(self, storage_path, version=''):
        safe = storage_path.lstrip('/').replace('\\', '/').replace('/', '__')
        tag = ('.' + str(version)[:10]) if version else ''
        return os.path.join(bg_cache_dir(), '_thumbs', safe + tag + '.jpg')

    def ensure_bg_thumb(self, storage_path, updated_at=None, size=(280, 158)):
        """Tải (nếu thiếu) rồi cắt thumbnail JPEG để gallery xem trước."""
        from PIL import Image
        thumb = self.bg_thumb_path(storage_path, updated_at or '')
        if os.path.exists(thumb):
            return {'ok': True, 'path': thumb}
        res = self.download_background(storage_path, updated_at)
        if not res.get('ok'):
            return res
        try:
            im = Image.open(res['path']).convert('RGB')
            im.thumbnail(size, Image.Resampling.LANCZOS)
            canvas = Image.new('RGB', size, (22, 22, 28))
            canvas.paste(im, ((size[0] - im.width) // 2, (size[1] - im.height) // 2))
            os.makedirs(os.path.dirname(thumb), exist_ok=True)
            canvas.save(thumb, 'JPEG', quality=82)
            return {'ok': True, 'path': thumb, 'full': res['path']}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    def download_background(self, storage_path, updated_at=None):
        local_fp = _local_path(bg_cache_dir(), storage_path.lstrip('/').replace('\\', '/'))
        meta_path = os.path.join(bg_cache_dir(), 'sync_meta.json')
        with self._api_lock:
            meta = _read_json(meta_path)
        if os.path.exists(local_fp) and meta.get('files', {}).get(storage_path) == updated_at:
            return {'ok': True, 'path': local_fp, 'fresh': False}
        if not self.logged_in():
            return {'ok': False, 'msg': 'Cần đăng nhập.'}
        try:
            self._request('GET', '/api/files/backgrounds/' + _quote_path(storage_path),
                          want='file', dest=local_fp, timeout=120)
        except CloudError as e:
            return {'ok': False, 'msg': e.msg}
        with self._api_lock:
            meta = _read_json(meta_path)
            meta.setdefault('files', {})[storage_path] = updated_at
            meta['synced_at'] = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
            _write_json(meta_path, meta)
        return {'ok': True, 'path': local_fp, 'fresh': True}

    def upload_background_files(self, paths, shared=False):
        """Đẩy 1 hoặc nhiều file nền. Nhân viên → users/{id}/; admin chọn shared/."""
        if not self.logged_in():
            return {'ok': False, 'msg': 'Cần đăng nhập.'}
        files = [p for p in paths if os.path.isfile(p) and p.lower().endswith(IMG_EXTS)]
        if not files:
            return {'ok': False, 'msg': 'Không có file ảnh hợp lệ.'}
        prefix = 'shared' if (shared and self.is_admin()) else f'users/{self.user_id}'
        ok = err = 0
        last_path = None
        last_err = ''
        for p in files:
            fn = _nfc(os.path.basename(p))
            sp = f'{prefix}/{fn}'
            try:
                with open(p, 'rb') as f:
                    data = f.read()
                self._request('PUT', '/api/files/backgrounds/' + _quote_path(sp), body=data,
                              headers={'Content-Type': 'application/octet-stream'}, timeout=180)
                ok += 1
                last_path = sp
            except (CloudError, OSError) as e:
                err += 1
                last_err = getattr(e, 'msg', str(e))
                print('BG upload err:', fn, last_err)
        return {'ok': err == 0, 'storage_path': last_path, 'prefix': prefix,
                'msg': f'Đã đẩy {ok}/{len(files)} ảnh nền lên máy chủ'
                       + (f', {err} lỗi ({last_err})' if err else '')}

    def upload_backgrounds(self, folder, shared=False):
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.lower().endswith(IMG_EXTS)]
        if not files:
            return {'ok': False, 'msg': 'Thư mục không có ảnh nền.'}
        return self.upload_background_files(files, shared=shared)

    # ── Tài khoản / quản trị ──
    def fetch_profiles(self):
        try:
            rows = self._request('GET', '/api/admin/users').get('users') or []
        except CloudError as e:
            print('Fetch profiles err:', e.msg)
            return []
        return [{'id': str(r.get('id')), 'name': r.get('name') or r.get('email'),
                 'email': r.get('email'), 'role': r.get('role')} for r in rows]

    def set_admin(self, uid, name, make_admin):
        if str(uid) == str(self.user_id) and not make_admin:
            return {'ok': False, 'msg': 'Không thể tự thu hồi quyền của chính mình.'}
        try:
            rows = self._request('GET', '/api/admin/users').get('users') or []
            u = next((r for r in rows if str(r.get('id')) == str(uid)), None)
            if not u:
                return {'ok': False, 'msg': 'Không có tài khoản này.'}
            self._request('POST', '/api/admin/users', body={
                'id': u['id'], 'email': u['email'], 'name': u.get('name') or name or '',
                'role': 'admin' if make_admin else 'user',
                'disabled': u.get('disabled', False)})
            return {'ok': True, 'msg': 'Đã cập nhật quyền.'}
        except CloudError as e:
            return {'ok': False, 'msg': e.msg}
