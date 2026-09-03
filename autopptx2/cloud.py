"""Supabase Auth + đồng bộ Excel / Avatar / ảnh nền.

Không import Tkinter — gọi từ worker thread. App chỉ nhận dict kết quả
rồi cập nhật UI trên luồng chính.
"""
import os
import re
import json
import base64
import datetime
import sys
from concurrent.futures import ThreadPoolExecutor

from .constants import (CONFIG_DIR, CONFIG_PATH, V1_CREDS_PATH, KEYRING_SERVICE,
                        SUPABASE_URL, SUPABASE_ANON_KEY, IMG_EXTS)

try:
    from supabase import create_client
except ImportError:
    create_client = None

try:
    import keyring
except ImportError:
    keyring = None

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
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def load_saved_creds():
    """Đọc email + mật khẩu đã nhớ (keyring, rồi fallback file V1/V2)."""
    email = pw = ''
    for path in (CONFIG_PATH, V1_CREDS_PATH):
        d = _read_json(path)
        if d.get('email'):
            email = d.get('email') or ''
            if d.get('pw'):
                try:
                    pw = base64.b64decode(d['pw'].encode()).decode('utf-8')
                except Exception:
                    pw = ''
            break
    if email and keyring:
        try:
            pw = keyring.get_password(KEYRING_SERVICE, email) or pw
        except Exception:
            pass
    return {'email': email, 'pw': pw} if email else None


def save_creds(email, pw):
    stored = False
    if keyring:
        try:
            keyring.set_password(KEYRING_SERVICE, email, pw)
            stored = True
        except Exception:
            pass
    d = _read_json(CONFIG_PATH)
    d['email'] = email
    d['remember'] = True
    if not stored:
        d['pw'] = base64.b64encode(pw.encode('utf-8')).decode()
    else:
        d.pop('pw', None)
    _write_json(CONFIG_PATH, d)


def clear_creds():
    d = _read_json(CONFIG_PATH)
    email = d.get('email', '')
    if keyring and email:
        try:
            keyring.delete_password(KEYRING_SERVICE, email)
        except Exception:
            pass
    d.pop('email', None)
    d.pop('pw', None)
    d.pop('remember', None)
    _write_json(CONFIG_PATH, d)


class CloudClient:
    def __init__(self, url=None, key=None):
        self.url = (url or SUPABASE_URL).strip()
        self.key = (key or SUPABASE_ANON_KEY).strip()
        self._client = None
        self.user_id = None
        self.email = None
        self.role = 'user'
        self.name = None

    def is_admin(self):
        return self.role in ('admin', 'super_admin')

    def logged_in(self):
        return bool(self.user_id)

    def package_ok(self):
        return create_client is not None

    def client(self):
        if create_client is None or not self.url or not self.key:
            return None
        if self._client is None:
            try:
                self._client = create_client(self.url, self.key)
            except Exception as e:
                print('Supabase client err:', e)
                self._client = None
        return self._client

    def _select_all(self, table, columns='*', page=1000):
        """Đọc hết hàng. PostgREST mặc định cắt 1000 — phải lật trang."""
        sb = self.client()
        if not sb:
            return []
        out, start = [], 0
        page = max(100, int(page or 1000))
        while True:
            res = (sb.table(table)
                   .select(columns)
                   .range(start, start + page - 1)
                   .execute())
            chunk = list(res.data or [])
            out.extend(chunk)
            if len(chunk) < page:
                break
            start += page
        return out

    # ── Auth ──
    def login(self, email, password):
        """Trả (ok, message). ok=True kể cả chế độ local-only (không có package)."""
        self.user_id = self.name = None
        self.role = 'user'
        self.email = email
        sb = self.client()
        if sb is None:
            if self.url and self.key and create_client is None:
                return False, 'Thiếu package supabase. Chạy: pip install supabase'
            return True, 'local'          # chưa cấu hình → vào máy này
        try:
            res = sb.auth.sign_in_with_password({'email': email, 'password': password})
            user = getattr(res, 'user', None)
            if user is None:
                return False, 'Sai email hoặc mật khẩu.'
            self.user_id = user.id
            try:
                prof = (sb.table('profiles').select('full_name, role')
                        .eq('id', user.id).single().execute())
                if prof.data:
                    self.role = prof.data.get('role') or 'user'
                    self.name = prof.data.get('full_name')
            except Exception as e:
                print('Load profile err:', e)
            return True, 'ok'
        except Exception as e:
            print('Supabase login err:', repr(e))
            return False, 'Sai email hoặc mật khẩu.'

    def logout(self):
        try:
            sb = self.client()
            if sb is not None:
                sb.auth.sign_out()
        except Exception:
            pass
        self._client = None
        self.user_id = self.email = self.name = None
        self.role = 'user'

    # ── Excel ──
    def excel_meta_path(self):
        return os.path.join(excel_cache_dir(), 'sync_meta.json')

    def excel_status_text(self):
        if not self.user_id:
            return 'Chưa đăng nhập'
        meta = _read_json(self.excel_meta_path()).get(self.user_id) or {}
        if not meta:
            return 'Chưa đồng bộ lần nào'
        return f"Lần cuối: {meta.get('synced_at_local', '')}"

    def sync_excel(self, progress=None):
        """Tải Excel của tài khoản đang đăng nhập. Trả dict kết quả."""
        sb = self.client()
        uid = self.user_id
        if not sb or not uid:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        if progress:
            progress(0.15, 'Đang kiểm tra bản Excel mới…')
        try:
            res = (sb.table('excel_data_files')
                   .select('storage_path, file_name, updated_at')
                   .eq('owner_id', uid).limit(1).execute())
            rows = res.data or []
        except Exception as e:
            return {'ok': False, 'msg': f'Lỗi kết nối: {e}'}
        if not rows:
            return {'ok': False, 'msg': 'Chưa có file Excel trên Cloud cho tài khoản này.'}
        row = rows[0]
        cache_path = os.path.join(excel_cache_dir(), f'{uid}.xlsx')
        meta = _read_json(self.excel_meta_path())
        local = meta.get(uid, {})
        if local.get('updated_at') == row.get('updated_at') and os.path.exists(cache_path):
            if progress:
                progress(1.0, 'Đã là bản mới nhất')
            return {'ok': True, 'path': cache_path, 'file_name': row.get('file_name'),
                    'fresh': False, 'msg': 'Excel đã là bản mới nhất.'}
        if progress:
            progress(0.6, 'Đang tải Excel…')
        try:
            data = sb.storage.from_('excel-data').download(row['storage_path'])
            with open(cache_path, 'wb') as f:
                f.write(data)
        except Exception as e:
            return {'ok': False, 'msg': f'Lỗi tải Excel: {e}'}
        meta[uid] = {
            'updated_at': row.get('updated_at'),
            'synced_at_local': datetime.datetime.now().strftime('%d/%m/%Y %H:%M'),
            'file_name': row.get('file_name') or os.path.basename(cache_path),
        }
        _write_json(self.excel_meta_path(), meta)
        if progress:
            progress(1.0, 'Đã cập nhật Excel')
        return {'ok': True, 'path': cache_path, 'file_name': row.get('file_name'),
                'fresh': True, 'msg': 'Đã tải Excel từ Cloud.'}

    def upload_excel(self, local_path, target_user_id=None):
        sb = self.client()
        uid = self.user_id
        if not sb or not uid:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        if not local_path or not os.path.exists(local_path):
            return {'ok': False, 'msg': 'Chưa chọn file Excel trên máy.'}
        target = target_user_id or uid
        storage_path = f'{target}/data.xlsx'
        try:
            with open(local_path, 'rb') as f:
                data = f.read()
            sb.storage.from_('excel-data').upload(
                storage_path, data,
                {'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                 'upsert': 'true'})
            now_iso = datetime.datetime.utcnow().isoformat()
            sb.table('excel_data_files').upsert({
                'owner_id': target,
                'storage_path': storage_path,
                'file_name': os.path.basename(local_path),
                'updated_at': now_iso,
                'updated_by': uid,
            }, on_conflict='owner_id').execute()
            if target == uid:
                meta = _read_json(self.excel_meta_path())
                meta[uid] = {
                    'updated_at': now_iso,
                    'synced_at_local': datetime.datetime.now().strftime('%d/%m/%Y %H:%M'),
                    'file_name': os.path.basename(local_path),
                }
                _write_json(self.excel_meta_path(), meta)
            return {'ok': True, 'msg': 'Đã đẩy Excel lên Cloud.'}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    def upload_excel_broadcast(self, local_path, profiles):
        """Admin: đẩy cùng 1 file cho mọi tài khoản trong danh sách profiles."""
        ok = err = 0
        last = ''
        for p in profiles:
            uid = p.get('id')
            if not uid:
                continue
            r = self.upload_excel(local_path, target_user_id=uid)
            if r.get('ok'):
                ok += 1
            else:
                err += 1
                last = r.get('msg', '')
        return {'ok': err == 0, 'msg': f'Đã đẩy cho {ok} tài khoản' + (f', {err} lỗi ({last})' if err else '')}

    # ── Avatar ──
    def avatar_meta_path(self):
        return os.path.join(avatar_cache_dir(), 'sync_meta.json')

    def avatar_status_text(self):
        d = avatar_cache_dir()
        n = 0
        if os.path.isdir(d):
            for _r, _ds, files in os.walk(d):
                n += sum(1 for f in files if f.lower().endswith(IMG_EXTS))
        last = _read_json(self.avatar_meta_path()).get('synced_at', '')
        return f'{n} ảnh trong kho' + (f' · {last}' if last else '')

    def sync_avatars(self, progress=None):
        sb = self.client()
        if not sb or not self.user_id:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        if progress:
            progress(0, 1, 'Đang kiểm tra kho avatar…')
        try:
            rows = self._select_all(
                'avatar_files', 'storage_path, file_name, updated_at')
        except Exception as e:
            return {'ok': False, 'msg': f'Lỗi kết nối kho avatar: {e}'}
        cache = avatar_cache_dir()
        meta = _read_json(self.avatar_meta_path())
        files_meta = meta.get('files', {})
        to_dl = []
        skipped = 0
        for r in rows:
            sp = r.get('storage_path')
            fn = r.get('file_name') or (os.path.basename(sp) if sp else '')
            up = r.get('updated_at')
            if not sp or not fn:
                continue
            rel = sp.lstrip('/').replace('\\', '/')
            local_fp = os.path.join(cache, *rel.split('/'))
            if os.path.exists(local_fp) and files_meta.get(sp) == up:
                skipped += 1
            else:
                to_dl.append((sp, up, local_fp))
        need = len(to_dl)
        downloaded = errors = 0
        if need:
            lock_n = {'n': 0, 'ok': 0, 'err': 0}

            def _one(item):
                sp, up, local_fp = item
                try:
                    data = sb.storage.from_('avatars').download(sp)
                    os.makedirs(os.path.dirname(local_fp), exist_ok=True)
                    with open(local_fp, 'wb') as f:
                        f.write(data)
                    files_meta[sp] = up
                    lock_n['ok'] += 1
                except Exception as e:
                    print('Avatar download err:', sp, e)
                    lock_n['err'] += 1
                lock_n['n'] += 1
                if progress and (lock_n['n'] % 4 == 0 or lock_n['n'] == need):
                    progress(lock_n['n'], need, f'Đang tải {lock_n["ok"]}/{need}…')

            with ThreadPoolExecutor(max_workers=10) as ex:
                list(ex.map(_one, to_dl))
            downloaded, errors = lock_n['ok'], lock_n['err']
        meta['files'] = files_meta
        meta['synced_at'] = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        _write_json(self.avatar_meta_path(), meta)
        msg = (f'Avatar: kho Cloud {len(rows)} · tải {downloaded} mới, '
               f'sẵn có {skipped}' + (f', {errors} lỗi' if errors else ''))
        if progress:
            progress(1, 1, msg)
        return {'ok': True, 'path': cache, 'downloaded': downloaded,
                'skipped': skipped, 'errors': errors, 'msg': msg}

    def upload_avatars(self, folder, progress=None):
        sb = self.client()
        if not sb or not self.user_id:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        files = []
        for root, _ds, fns in os.walk(folder):
            for fn in fns:
                if fn.lower().endswith(IMG_EXTS):
                    files.append(os.path.join(root, fn))
        if not files:
            return {'ok': False, 'msg': 'Thư mục không có ảnh avatar.'}
        ctype = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                 '.webp': 'image/webp'}
        ok = err = 0
        for i, path in enumerate(files, 1):
            fn = os.path.basename(path)
            try:
                with open(path, 'rb') as f:
                    data = f.read()
                rel = os.path.relpath(path, folder).replace('\\', '/')
                sp = f'shared/{rel}'
                ext = os.path.splitext(fn)[1].lower()
                sb.storage.from_('avatars').upload(
                    sp, data, {'content-type': ctype.get(ext, 'image/png'), 'upsert': 'true'})
                code = _AVATAR_STRIP.sub('', os.path.splitext(fn)[0])
                sb.table('avatar_files').upsert({
                    'storage_path': sp, 'file_name': fn, 'code': code,
                    'size': len(data),
                    'updated_at': datetime.datetime.utcnow().isoformat(),
                }, on_conflict='storage_path').execute()
                ok += 1
            except Exception as e:
                err += 1
                print('Avatar upload err:', fn, e)
            if progress:
                progress(i, len(files), f'Đẩy {i}/{len(files)}…')
        return {'ok': err == 0,
                'msg': f'Đã đẩy {ok}/{len(files)} avatar' + (f', {err} lỗi' if err else '')}

    # ── Ảnh nền ──
    def list_backgrounds(self, prefix=''):
        sb = self.client()
        if not sb:
            return []
        out = []
        try:
            items = sb.storage.from_('backgrounds').list(prefix) or []
        except Exception as e:
            print('BG storage list err:', prefix, e)
            return out
        for it in items:
            name = it.get('name') if isinstance(it, dict) else getattr(it, 'name', None)
            if not name or name.startswith('.'):
                continue
            full = f'{prefix}/{name}' if prefix else name
            meta = it.get('metadata') if isinstance(it, dict) else getattr(it, 'metadata', None)
            _id = it.get('id') if isinstance(it, dict) else getattr(it, 'id', None)
            if _id is None and not meta:
                out.extend(self.list_backgrounds(full))
            elif name.lower().endswith(IMG_EXTS):
                up = (it.get('updated_at') if isinstance(it, dict)
                      else getattr(it, 'updated_at', None)) or ''
                out.append({'storage_path': full, 'file_name': name, 'updated_at': up})
        return out

    def bg_thumb_path(self, storage_path):
        safe = storage_path.lstrip('/').replace('\\', '/').replace('/', '__')
        return os.path.join(bg_cache_dir(), '_thumbs', safe + '.jpg')

    def ensure_bg_thumb(self, storage_path, updated_at=None, size=(280, 158)):
        """Tải (nếu thiếu) rồi cắt thumbnail JPEG để gallery xem trước."""
        from PIL import Image
        thumb = self.bg_thumb_path(storage_path)
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
        local_fp = os.path.join(bg_cache_dir(), *storage_path.lstrip('/').replace('\\', '/').split('/'))
        meta_path = os.path.join(bg_cache_dir(), 'sync_meta.json')
        meta = _read_json(meta_path)
        if os.path.exists(local_fp) and meta.get('files', {}).get(storage_path) == updated_at:
            return {'ok': True, 'path': local_fp, 'fresh': False}
        sb = self.client()
        if not sb:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        try:
            data = sb.storage.from_('backgrounds').download(storage_path)
            os.makedirs(os.path.dirname(local_fp), exist_ok=True)
            with open(local_fp, 'wb') as f:
                f.write(data)
            meta.setdefault('files', {})[storage_path] = updated_at
            meta['synced_at'] = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
            _write_json(meta_path, meta)
            return {'ok': True, 'path': local_fp, 'fresh': True}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    def upload_background_files(self, paths, shared=False):
        """Đẩy 1 hoặc nhiều file nền. Nhân viên → users/{uid}/; admin chọn shared/."""
        sb = self.client()
        if not sb or not self.user_id:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        files = [p for p in paths if os.path.isfile(p) and p.lower().endswith(IMG_EXTS)]
        if not files:
            return {'ok': False, 'msg': 'Không có file ảnh hợp lệ.'}
        prefix = 'shared' if (shared and self.is_admin()) else f'users/{self.user_id}'
        ctype = {'.png': 'image/png', '.jpg': 'image/jpeg',
                 '.jpeg': 'image/jpeg', '.webp': 'image/webp'}
        ok = err = 0
        last_path = None
        for p in files:
            fn = os.path.basename(p)
            try:
                with open(p, 'rb') as f:
                    data = f.read()
                ext = os.path.splitext(fn)[1].lower()
                sp = f'{prefix}/{fn}'
                sb.storage.from_('backgrounds').upload(
                    sp, data, {'content-type': ctype.get(ext, 'image/png'), 'upsert': 'true'})
                ok += 1
                last_path = sp
            except Exception as e:
                err += 1
                print('BG upload err:', fn, e)
        return {'ok': err == 0, 'storage_path': last_path, 'prefix': prefix,
                'msg': f'Đã đẩy {ok}/{len(files)} ảnh nền lên Cloud'
                       + (f', {err} lỗi' if err else '')}

    def upload_backgrounds(self, folder, shared=False):
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.lower().endswith(IMG_EXTS)]
        if not files:
            return {'ok': False, 'msg': 'Thư mục không có ảnh nền.'}
        return self.upload_background_files(files, shared=shared)

    # ── Profiles / admin ──
    def fetch_profiles(self):
        sb = self.client()
        if not sb:
            return []
        try:
            rows = self._select_all('profiles', 'id, full_name, role')
            rows.sort(key=lambda r: str(r.get('full_name') or '').casefold())
            return [{'id': r.get('id'), 'name': r.get('full_name'), 'role': r.get('role')}
                    for r in rows]
        except Exception as e:
            print('Fetch profiles err:', e)
            return []

    def set_admin(self, uid, name, make_admin):
        sb = self.client()
        if not sb:
            return {'ok': False, 'msg': 'Cần đăng nhập Cloud.'}
        try:
            if make_admin:
                sb.table('admins').upsert(
                    {'id': uid, 'full_name': name or 'Quản trị viên'},
                    on_conflict='id').execute()
            else:
                if uid == self.user_id:
                    return {'ok': False, 'msg': 'Không thể tự thu hồi quyền của chính mình.'}
                sb.table('admins').delete().eq('id', uid).execute()
            return {'ok': True, 'msg': 'Đã cập nhật quyền.'}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}
