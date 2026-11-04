#!/usr/bin/env python3
"""AutoPPTX Studio — máy chủ kho tài nguyên (thay Supabase).

Lưu FILE TỔNG Excel, kho avatar (ảnh theo mã RP), ảnh nền và tài khoản người dùng
trên VPS của công ty. Chỉ dùng thư viện chuẩn Python; Pillow (thu nhỏ avatar) và
openpyxl (đối chiếu mã RP ↔ avatar) là tuỳ chọn.

Chạy:
    python3 autopptx_server.py                      # phục vụ API (mặc định)
    python3 autopptx_server.py reindex              # quét lại thư mục store/ vào DB
    python3 autopptx_server.py create-admin EMAIL [TÊN]   # in mật khẩu tạm 1 lần
    python3 autopptx_server.py set-password EMAIL   # in mật khẩu tạm mới

Biến môi trường:
    AP_DATA          thư mục dữ liệu (mặc định ./data)
    AP_BIND, AP_PORT nơi lắng nghe (127.0.0.1:8780)
    AP_ACCEL_PREFIX  nếu đặt (vd /pptx-store/), file trả qua nginx X-Accel-Redirect
    AP_URL_PREFIX    tiền tố đường dẫn phía nginx (vd /pptx) — máy chủ tự bỏ đi
    AP_COOKIE_PATH   Path cookie phiên cho trang quản trị (mặc định /)
    AP_COOKIE_SECURE 1 = cookie Secure (bật khi chạy sau HTTPS)
    AP_AVATAR_MAX_PX cạnh dài tối đa của avatar khi lưu (1600; 0 = giữ nguyên)
    AP_MAX_UPLOAD_MB giới hạn 1 file tải lên (64)

Đường dẫn API (nginx chuyển /pptx/… → /…):
    GET  /                       trang quản trị web
    GET  /api/health
    POST /api/login              {email, password} → {token, user}
    POST /api/logout
    GET  /api/me
    POST /api/me/password        {old, new}
    GET  /api/manifest/avatars   danh sách avatar (path, size, sha, src_sha, code, updated_at)
    GET  /api/manifest/backgrounds   nền chung + nền riêng của mình
    GET  /api/files/<kind>/<path>    tải file
    PUT  /api/files/<kind>/<path>    đẩy file (avatar: admin; nền: shared/ admin, users/<id>/ chủ)
    DELETE /api/files/<kind>/<path>
    GET  /api/excel/meta         FILE TỔNG đang áp cho mình (riêng nếu có, không thì chung)
    GET  /api/excel/file
    PUT  /api/excel?target=me|all|<id>   header X-File-Name (percent-encoded)
    GET  /api/admin/users        POST /api/admin/users      DELETE /api/admin/users/<id>
    POST /api/admin/users/<id>/password   → mật khẩu tạm mới
    GET  /api/admin/stats
    GET  /api/admin/report[?owner=shared|<id>]   đối chiếu mã RP ↔ avatar
"""
import datetime
import hashlib
import hmac
import io
import json
import os
import re
import secrets
import shutil
import sqlite3
import sys
import threading
import time
import unicodedata
import urllib.parse
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.abspath(os.environ.get('AP_DATA') or os.path.join(HERE, 'data'))
STORE = os.path.join(DATA, 'store')
DB_PATH = os.path.join(DATA, 'autopptx.db')
BIND = os.environ.get('AP_BIND', '127.0.0.1')
PORT = int(os.environ.get('AP_PORT', '8780'))
ACCEL_PREFIX = os.environ.get('AP_ACCEL_PREFIX', '').strip()
URL_PREFIX = os.environ.get('AP_URL_PREFIX', '').strip().rstrip('/')
COOKIE_PATH = os.environ.get('AP_COOKIE_PATH', '/') or '/'
COOKIE_SECURE = os.environ.get('AP_COOKIE_SECURE', '') == '1'
AVATAR_MAX_PX = int(os.environ.get('AP_AVATAR_MAX_PX', '1600') or 0)
MAX_UPLOAD = int(os.environ.get('AP_MAX_UPLOAD_MB', '64')) * 1024 * 1024

KINDS = ('avatars', 'backgrounds')
IMG_EXTS = ('.png', '.jpg', '.jpeg', '.webp')
ROLES = ('user', 'admin', 'super_admin')
SESSION_DAYS = 60
COOKIE_NAME = 'ap_session'
PBKDF2_ITER = 240_000
# Cùng quy tắc với app (cloud.py / datasource.py): bỏ hậu tố A, AA, _Ava → mã RP.
AVATAR_SUFFIX = re.compile(r'(_Ava|AA|A)$', re.IGNORECASE)

CTYPES = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
          '.webp': 'image/webp',
          '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}

try:
    from PIL import Image, ImageOps
except Exception:          # Pillow là tuỳ chọn: thiếu thì lưu nguyên ảnh.
    Image = ImageOps = None


# ═══════════════════════════ tiện ích ═══════════════════════════
def now_iso():
    return datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()


def sha256_bytes(b):
    return hashlib.sha256(b).hexdigest()


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def hash_password(pw):
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac('sha256', pw.encode('utf-8'), salt, PBKDF2_ITER)
    return f'pbkdf2${PBKDF2_ITER}${salt.hex()}${dk.hex()}'


def check_password(pw, stored):
    try:
        algo, it, salt, want = (stored or '').split('$')
        if algo != 'pbkdf2':
            return False
        dk = hashlib.pbkdf2_hmac('sha256', (pw or '').encode('utf-8'),
                                 bytes.fromhex(salt), int(it))
        return hmac.compare_digest(dk.hex(), want)
    except Exception:
        return False


def new_temp_password():
    # Dễ đọc qua điện thoại: bỏ ký tự dễ nhầm (0/O, 1/l/I).
    alphabet = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789'
    return ''.join(secrets.choice(alphabet) for _ in range(12))


def avatar_code(rel_path):
    stem = os.path.splitext(os.path.basename(rel_path))[0].strip().upper()
    return AVATAR_SUFFIX.sub('', stem).strip() or stem


def clean_rel(path, kind):
    """Chuẩn hoá đường dẫn tương đối; None nếu không an toàn / sai loại file."""
    p = unicodedata.normalize('NFC', str(path or '')).replace('\\', '/').strip('/')
    parts = [s.strip() for s in p.split('/')]
    if not parts or any(s in ('', '.', '..') for s in parts):
        return None
    if any(ch in s for s in parts for ch in '\x00:*?"<>|') or any(s.startswith('.') for s in parts):
        return None
    if len(parts) > 8 or len(p) > 400:
        return None
    if not parts[-1].lower().endswith(IMG_EXTS):
        return None
    if kind == 'backgrounds':
        if parts[0] == 'shared' and len(parts) >= 2:
            pass
        elif parts[0] == 'users' and len(parts) >= 3 and parts[1].isdigit():
            pass
        else:
            return None
    return '/'.join(parts)


def store_path(kind, rel):
    return os.path.join(STORE, kind, *rel.split('/'))


def atomic_write(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = f'{path}.{secrets.token_hex(4)}.part'
    with open(tmp, 'wb') as f:
        f.write(data)
    os.replace(tmp, path)


def shrink_avatar(data, ext):
    """Thu nhỏ avatar quá lớn (giữ định dạng). Trả bytes gốc nếu không lợi hơn."""
    if not AVATAR_MAX_PX or Image is None:
        return data
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
        if max(im.size) <= AVATAR_MAX_PX and len(data) <= 900_000:
            return data
        im = ImageOps.exif_transpose(im)
        im.thumbnail((AVATAR_MAX_PX, AVATAR_MAX_PX), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        e = ext.lower()
        if e in ('.jpg', '.jpeg'):
            if im.mode not in ('RGB', 'L'):
                im = im.convert('RGB')
            im.save(out, 'JPEG', quality=85, optimize=True, progressive=True)
        elif e == '.png':
            im.save(out, 'PNG', optimize=True)
        elif e == '.webp':
            im.save(out, 'WEBP', quality=85)
        else:
            return data
        b = out.getvalue()
        return b if 0 < len(b) < len(data) else data
    except Exception:
        return data


# ═══════════════════════════ cơ sở dữ liệu ═══════════════════════════
SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL UNIQUE COLLATE NOCASE,
    name TEXT NOT NULL DEFAULT '',
    role TEXT NOT NULL DEFAULT 'user',
    pass_hash TEXT NOT NULL DEFAULT '',
    must_change INTEGER NOT NULL DEFAULT 1,
    disabled INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    last_login TEXT
);
CREATE TABLE IF NOT EXISTS sessions (
    token_hash TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    created_at TEXT NOT NULL,
    last_seen REAL NOT NULL,
    expires REAL NOT NULL,
    client TEXT
);
CREATE TABLE IF NOT EXISTS files (
    kind TEXT NOT NULL,
    path TEXT NOT NULL,
    size INTEGER NOT NULL,
    sha TEXT NOT NULL,
    src_sha TEXT NOT NULL,
    code TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL,
    updated_by INTEGER,
    PRIMARY KEY (kind, path)
);
CREATE TABLE IF NOT EXISTS excel (
    owner TEXT PRIMARY KEY,
    file_name TEXT NOT NULL,
    size INTEGER NOT NULL,
    sha TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    updated_by INTEGER
);
CREATE TABLE IF NOT EXISTS audit (
    ts TEXT NOT NULL,
    user_id INTEGER,
    action TEXT NOT NULL,
    detail TEXT
);
"""

_write_lock = threading.Lock()


def db():
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys=ON')
    return con


def init_db():
    os.makedirs(DATA, exist_ok=True)
    for k in KINDS + ('excel',):
        os.makedirs(os.path.join(STORE, k), exist_ok=True)
    con = db()
    con.execute('PRAGMA journal_mode=WAL')
    con.executescript(SCHEMA)
    con.commit()
    con.close()


def audit(con, uid, action, detail=''):
    con.execute('INSERT INTO audit (ts, user_id, action, detail) VALUES (?,?,?,?)',
                (now_iso(), uid, action, str(detail)[:500]))


def user_public(u):
    return {'id': u['id'], 'email': u['email'], 'name': u['name'], 'role': u['role'],
            'must_change': bool(u['must_change']), 'disabled': bool(u['disabled']),
            'last_login': u['last_login'], 'created_at': u['created_at']}


def record_file(con, kind, rel, data, src_sha, uid):
    con.execute(
        'INSERT INTO files (kind, path, size, sha, src_sha, code, updated_at, updated_by) '
        'VALUES (?,?,?,?,?,?,?,?) ON CONFLICT(kind, path) DO UPDATE SET '
        'size=excluded.size, sha=excluded.sha, src_sha=excluded.src_sha, code=excluded.code, '
        'updated_at=excluded.updated_at, updated_by=excluded.updated_by',
        (kind, rel, len(data), sha256_bytes(data), src_sha,
         avatar_code(rel) if kind == 'avatars' else '', now_iso(), uid))


def reindex():
    """Đồng bộ bảng files với thư mục store/ (sau khi chép hàng loạt bằng scp/rsync)."""
    init_db()
    con = db()
    added = updated = removed = 0
    seen = set()
    with _write_lock:
        for kind in KINDS:
            base = os.path.join(STORE, kind)
            have = {r['path']: r for r in con.execute(
                'SELECT path, size, sha FROM files WHERE kind=?', (kind,))}
            for root, _ds, fns in os.walk(base):
                for fn in fns:
                    full = os.path.join(root, fn)
                    rel_raw = os.path.relpath(full, base).replace(os.sep, '/')
                    rel = clean_rel(rel_raw, kind)
                    if not rel:
                        continue
                    if rel != rel_raw:          # tên NFD (macOS) → NFC
                        dst = store_path(kind, rel)
                        os.makedirs(os.path.dirname(dst), exist_ok=True)
                        os.replace(full, dst)
                        full = dst
                    seen.add((kind, rel))
                    size = os.path.getsize(full)
                    old = have.get(rel)
                    if old and old['size'] == size:
                        continue
                    sha = sha256_file(full)
                    if old and old['sha'] == sha:
                        continue
                    con.execute(
                        'INSERT INTO files (kind, path, size, sha, src_sha, code, updated_at) '
                        'VALUES (?,?,?,?,?,?,?) ON CONFLICT(kind, path) DO UPDATE SET '
                        'size=excluded.size, sha=excluded.sha, src_sha=excluded.src_sha, '
                        'code=excluded.code, updated_at=excluded.updated_at',
                        (kind, rel, size, sha, sha,
                         avatar_code(rel) if kind == 'avatars' else '', now_iso()))
                    if old:
                        updated += 1
                    else:
                        added += 1
            for rel in have:
                if (kind, rel) not in seen:
                    con.execute('DELETE FROM files WHERE kind=? AND path=?', (kind, rel))
                    removed += 1
        con.commit()
    con.close()
    return {'added': added, 'updated': updated, 'removed': removed}


def apply_src_map(map_path):
    """Sau khi chép hàng loạt ảnh ĐÃ thu nhỏ: ghi sha của ảnh gốc ({kind: {path: sha}})
    để lần đẩy lại từ thư mục gốc biết ảnh nào không đổi."""
    with open(map_path, 'r', encoding='utf-8') as f:
        m = json.load(f)
    con = db()
    n = 0
    with _write_lock:
        for kind, paths in m.items():
            if kind not in KINDS:
                continue
            for rel, sha in paths.items():
                cur = con.execute('UPDATE files SET src_sha=? WHERE kind=? AND path=?',
                                  (sha, kind, unicodedata.normalize('NFC', rel)))
                n += cur.rowcount
        con.commit()
    con.close()
    return n


# ═══════════════════════════ đối chiếu mã RP ═══════════════════════════
_report_cache = {}
_report_lock = threading.Lock()


def build_report(excel_path, excel_sha):
    """Dùng chính ExcelSource/AvatarIndex của app để kết quả trùng với lúc xuất PPTX."""
    con = db()
    av_ver = con.execute(
        "SELECT COUNT(*) AS n, COALESCE(MAX(updated_at),'') AS t FROM files WHERE kind='avatars'"
    ).fetchone()
    con.close()
    key = (excel_sha, av_ver['n'], av_ver['t'])
    with _report_lock:
        if key in _report_cache:
            return _report_cache[key]
    try:
        # Trên VPS: bản sao autopptx2/ nằm cạnh file này. Chạy từ repo: ở thư mục cha.
        for d in (HERE, os.path.dirname(HERE)):
            if os.path.isdir(os.path.join(d, 'autopptx2')) and d not in sys.path:
                sys.path.insert(0, d)
        from autopptx2.datasource import ExcelSource, AvatarIndex   # noqa: E402
    except Exception as e:
        return {'ok': False, 'msg': f'Máy chủ chưa có module đối chiếu: {e}'}
    src = ExcelSource()
    try:
        ok = src.load(excel_path)
    except Exception as e:
        return {'ok': False, 'msg': f'Không đọc được FILE TỔNG: {e}'}
    if not ok and not getattr(src, 'rows', None):
        return {'ok': False, 'msg': getattr(src, 'error', '') or 'Không đọc được FILE TỔNG.'}
    idx = AvatarIndex()
    idx.set_folder(os.path.join(STORE, 'avatars'))
    codes, seen = [], set()
    for r in src.rows:
        c = str(r.get('Code_RP') or '').strip().upper()
        if c and c not in seen:
            seen.add(c)
            codes.append((c, str(r.get('Name') or '').strip(), str(r.get('Channel') or '').strip()))
    base = os.path.join(STORE, 'avatars')
    used, missing, matched = set(), [], 0
    for c, name, ch in codes:
        p = idx.find(c)
        if p:
            matched += 1
            used.add(os.path.relpath(p, base).replace(os.sep, '/'))
        else:
            missing.append({'code': c, 'name': name, 'channel': ch})
    con = db()
    rows = con.execute("SELECT path, code FROM files WHERE kind='avatars' ORDER BY path").fetchall()
    con.close()
    unused = [r['path'] for r in rows if r['path'] not in used]
    by_code = {}
    for r in rows:
        by_code.setdefault(r['code'], []).append(r['path'])
    dups = [{'code': c, 'paths': ps} for c, ps in sorted(by_code.items()) if len(ps) > 1]
    res = {'ok': True, 'codes': len(codes), 'matched': matched, 'missing': missing,
           'avatars': len(rows), 'unused': unused, 'duplicates': dups,
           'generated_at': now_iso()}
    with _report_lock:
        _report_cache.clear()
        _report_cache[key] = res
    return res


# ═══════════════════════════ chặn dò mật khẩu ═══════════════════════════
_fail_lock = threading.Lock()
_fails = {}           # email → [thời điểm sai]


def login_blocked(email):
    t = time.time()
    with _fail_lock:
        lst = [x for x in _fails.get(email, []) if t - x < 600]
        _fails[email] = lst
        return len(lst) >= 6


def login_failed(email):
    with _fail_lock:
        _fails.setdefault(email, []).append(time.time())


# ═══════════════════════════ HTTP ═══════════════════════════
class ApiError(Exception):
    def __init__(self, status, msg):
        super().__init__(msg)
        self.status = status
        self.msg = msg


class Handler(BaseHTTPRequestHandler):
    server_version = 'AutoPPTXStore/1.0'
    protocol_version = 'HTTP/1.1'

    def log_message(self, fmt, *args):
        sys.stderr.write('%s %s\n' % (self.headers.get('X-Real-IP') or self.client_address[0],
                                      fmt % args))

    # ── phản hồi ──
    def send_json(self, obj, status=200, headers=None):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)

    def send_file(self, kind_dir, rel, ctype, download_name=None):
        full = os.path.join(STORE, kind_dir, *rel.split('/'))
        if not os.path.isfile(full):
            raise ApiError(404, 'Không có file.')
        if ACCEL_PREFIX:
            uri = ACCEL_PREFIX.rstrip('/') + '/' + kind_dir + '/' + urllib.parse.quote(rel)
            self.send_response(200)
            self.send_header('X-Accel-Redirect', uri)
            self.send_header('Content-Type', ctype)
            if download_name:
                self.send_header('Content-Disposition', "attachment; filename*=UTF-8''"
                                 + urllib.parse.quote(download_name))
            self.send_header('Content-Length', '0')
            self.end_headers()
            return
        size = os.path.getsize(full)
        self.send_response(200)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(size))
        self.send_header('Cache-Control', 'private, no-cache')
        if download_name:
            self.send_header('Content-Disposition', "attachment; filename*=UTF-8''"
                             + urllib.parse.quote(download_name))
        self.end_headers()
        if self.command == 'HEAD':
            return
        with open(full, 'rb') as f:
            shutil.copyfileobj(f, self.wfile, 1 << 16)

    # ── đọc yêu cầu ──
    def read_body(self, limit=MAX_UPLOAD):
        n = int(self.headers.get('Content-Length') or 0)
        if n > limit:
            raise ApiError(413, f'File quá lớn (giới hạn {limit // (1024 * 1024)} MB).')
        buf = bytearray()
        while len(buf) < n:
            chunk = self.rfile.read(min(1 << 20, n - len(buf)))
            if not chunk:
                break
            buf += chunk
        if len(buf) != n:
            raise ApiError(400, 'Tải lên bị ngắt giữa chừng.')
        return bytes(buf)

    def read_json(self):
        raw = self.read_body(limit=64 * 1024)
        try:
            return json.loads(raw.decode('utf-8') or '{}')
        except Exception:
            raise ApiError(400, 'Dữ liệu gửi lên không hợp lệ.')

    def token(self):
        auth = self.headers.get('Authorization') or ''
        if auth.lower().startswith('bearer '):
            return auth[7:].strip(), False
        for part in (self.headers.get('Cookie') or '').split(';'):
            k, _, v = part.strip().partition('=')
            if k == COOKIE_NAME and v:
                return v, True
        return '', False

    def current_user(self, con, required=True):
        tok, via_cookie = self.token()
        u = None
        if tok:
            th = sha256_bytes(tok.encode())
            row = con.execute(
                'SELECT s.last_seen, s.expires, u.* FROM sessions s JOIN users u ON u.id=s.user_id '
                'WHERE s.token_hash=?', (th,)).fetchone()
            t = time.time()
            if row and row['expires'] > t and not row['disabled']:
                u = row
                if t - row['last_seen'] > 3600:      # gia hạn trượt, ghi tối đa 1 lần/giờ
                    with _write_lock:
                        con.execute('UPDATE sessions SET last_seen=?, expires=? WHERE token_hash=?',
                                    (t, t + SESSION_DAYS * 86400, th))
                        con.commit()
        if u is None:
            if required:
                raise ApiError(401, 'Phiên đăng nhập đã hết. Đăng nhập lại.')
            return None
        # Cookie chỉ dùng cho trang quản trị; yêu cầu ghi phải kèm header riêng (chống CSRF).
        if via_cookie and self.command not in ('GET', 'HEAD') \
                and self.headers.get('X-Requested-With') != 'autopptx':
            raise ApiError(403, 'Thiếu header bảo vệ.')
        return u

    @staticmethod
    def need_admin(u):
        if u['role'] not in ('admin', 'super_admin'):
            raise ApiError(403, 'Chỉ quản trị viên mới làm được việc này.')

    # ── định tuyến ──
    def do_GET(self):
        self.route()

    def do_HEAD(self):
        self.route()

    def do_POST(self):
        self.route()

    def do_PUT(self):
        self.route()

    def do_DELETE(self):
        self.route()

    def route(self):
        raw_path = self.path
        # nginx chuyển nguyên URI (không viết lại, để giữ đúng tên file tiếng Việt);
        # bỏ tiền tố /pptx ở đây.
        if URL_PREFIX and (raw_path == URL_PREFIX or raw_path.startswith(URL_PREFIX + '/')
                           or raw_path.startswith(URL_PREFIX + '?')):
            raw_path = raw_path[len(URL_PREFIX):] or '/'
            if raw_path.startswith('?'):
                raw_path = '/' + raw_path
        url = urllib.parse.urlsplit(raw_path)
        path = urllib.parse.unquote(url.path)
        q = dict(urllib.parse.parse_qsl(url.query))
        m = self.command
        con = db()
        try:
            if path in ('/', '/index.html') and m in ('GET', 'HEAD'):
                return self.serve_ui()
            if not path.startswith('/api/'):
                raise ApiError(404, 'Không có trang này.')
            p = path[4:]
            if p == '/health':
                return self.send_json({'ok': True, 'time': now_iso()})
            if p == '/login' and m == 'POST':
                return self.api_login(con)
            if p == '/logout' and m == 'POST':
                return self.api_logout(con)
            u = self.current_user(con)
            if p == '/me' and m == 'GET':
                return self.send_json({'user': user_public(u)})
            if p == '/me/password' and m == 'POST':
                return self.api_change_password(con, u)
            mm = re.match(r'^/manifest/(avatars|backgrounds)$', p)
            if mm and m == 'GET':
                return self.api_manifest(con, u, mm.group(1))
            mm = re.match(r'^/files/(avatars|backgrounds)/(.+)$', p)
            if mm:
                return self.api_file(con, u, mm.group(1), mm.group(2))
            if p == '/excel/meta' and m == 'GET':
                return self.api_excel_meta(con, u)
            if p == '/excel/file' and m in ('GET', 'HEAD'):
                return self.api_excel_file(con, u, q)
            if p == '/excel' and m == 'PUT':
                return self.api_excel_put(con, u, q)
            if p.startswith('/admin/'):
                self.need_admin(u)
                return self.api_admin(con, u, p[6:], q)
            raise ApiError(404, 'Không có API này.')
        except ApiError as e:
            self.send_json({'ok': False, 'msg': e.msg}, status=e.status)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:      # lỗi lạ: ghi log, không lộ chi tiết
            import traceback
            traceback.print_exc()
            try:
                self.send_json({'ok': False, 'msg': f'Lỗi máy chủ: {type(e).__name__}'}, status=500)
            except Exception:
                pass
        finally:
            con.close()

    # ── trang quản trị ──
    def serve_ui(self):
        fp = os.path.join(HERE, 'admin.html')
        try:
            with open(fp, 'rb') as f:
                body = f.read()
        except OSError:
            raise ApiError(404, 'Thiếu admin.html')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('X-Frame-Options', 'DENY')
        self.send_header('X-Robots-Tag', 'noindex')
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(body)

    # ── đăng nhập ──
    def api_login(self, con):
        d = self.read_json()
        email = str(d.get('email') or '').strip().lower()
        pw = str(d.get('password') or '')
        client = str(d.get('client') or self.headers.get('User-Agent') or '')[:120]
        if not email or not pw:
            raise ApiError(400, 'Nhập email và mật khẩu.')
        if login_blocked(email):
            raise ApiError(429, 'Sai quá nhiều lần. Thử lại sau 10 phút.')
        u = con.execute('SELECT * FROM users WHERE email=?', (email,)).fetchone()
        if not u or u['disabled'] or not check_password(pw, u['pass_hash']):
            login_failed(email)
            time.sleep(0.8)
            raise ApiError(401, 'Sai email hoặc mật khẩu.')
        tok = secrets.token_urlsafe(32)
        t = time.time()
        with _write_lock:
            con.execute('DELETE FROM sessions WHERE expires < ?', (t,))
            con.execute('INSERT INTO sessions (token_hash, user_id, created_at, last_seen, expires, client) '
                        'VALUES (?,?,?,?,?,?)',
                        (sha256_bytes(tok.encode()), u['id'], now_iso(), t,
                         t + SESSION_DAYS * 86400, client))
            con.execute('UPDATE users SET last_login=? WHERE id=?', (now_iso(), u['id']))
            audit(con, u['id'], 'login', client)
            con.commit()
        headers = {}
        if d.get('cookie'):
            c = (f'{COOKIE_NAME}={tok}; Path={COOKIE_PATH}; HttpOnly; SameSite=Strict; '
                 f'Max-Age={SESSION_DAYS * 86400}')
            if COOKIE_SECURE:
                c += '; Secure'
            headers['Set-Cookie'] = c
        self.send_json({'ok': True, 'token': None if d.get('cookie') else tok,
                        'user': user_public(u)}, headers=headers)

    def api_logout(self, con):
        tok, _ = self.token()
        if tok:
            with _write_lock:
                con.execute('DELETE FROM sessions WHERE token_hash=?', (sha256_bytes(tok.encode()),))
                con.commit()
        c = f'{COOKIE_NAME}=; Path={COOKIE_PATH}; HttpOnly; SameSite=Strict; Max-Age=0'
        self.send_json({'ok': True}, headers={'Set-Cookie': c})

    def api_change_password(self, con, u):
        d = self.read_json()
        old, new = str(d.get('old') or ''), str(d.get('new') or '')
        if not check_password(old, u['pass_hash']):
            time.sleep(0.8)
            raise ApiError(400, 'Mật khẩu hiện tại không đúng.')
        if len(new) < 8:
            raise ApiError(400, 'Mật khẩu mới cần ít nhất 8 ký tự.')
        with _write_lock:
            con.execute('UPDATE users SET pass_hash=?, must_change=0 WHERE id=?',
                        (hash_password(new), u['id']))
            tok, _ = self.token()
            con.execute('DELETE FROM sessions WHERE user_id=? AND token_hash<>?',
                        (u['id'], sha256_bytes(tok.encode())))
            audit(con, u['id'], 'password')
            con.commit()
        self.send_json({'ok': True, 'msg': 'Đã đổi mật khẩu.'})

    # ── kho file ──
    def api_manifest(self, con, u, kind):
        if kind == 'backgrounds':
            rows = con.execute(
                "SELECT * FROM files WHERE kind='backgrounds' AND (path LIKE 'shared/%' OR path LIKE ?) "
                'ORDER BY path', (f'users/{u["id"]}/%',)).fetchall()
        else:
            rows = con.execute("SELECT * FROM files WHERE kind='avatars' ORDER BY path").fetchall()
        files = [{'path': r['path'], 'size': r['size'], 'sha': r['sha'], 'src_sha': r['src_sha'],
                  'code': r['code'], 'updated_at': r['updated_at']} for r in rows]
        self.send_json({'ok': True, 'kind': kind, 'count': len(files),
                        'bytes': sum(f['size'] for f in files), 'files': files})

    def can_write_bg(self, u, rel):
        if rel.startswith('shared/'):
            return u['role'] in ('admin', 'super_admin')
        return rel.startswith(f'users/{u["id"]}/') or u['role'] in ('admin', 'super_admin')

    def api_file(self, con, u, kind, raw):
        rel = clean_rel(raw, kind)
        if not rel:
            raise ApiError(400, 'Tên/đường dẫn file không hợp lệ (chỉ ảnh png/jpg/jpeg/webp).')
        m = self.command
        if m in ('GET', 'HEAD'):
            if kind == 'backgrounds' and rel.startswith('users/') \
                    and not self.can_write_bg(u, rel):
                raise ApiError(403, 'Không xem được ảnh nền riêng của người khác.')
            ext = os.path.splitext(rel)[1].lower()
            return self.send_file(kind, rel, CTYPES.get(ext, 'application/octet-stream'))
        if m == 'PUT':
            if kind == 'avatars':
                self.need_admin(u)
            elif not self.can_write_bg(u, rel):
                raise ApiError(403, 'Chỉ quản trị viên được đẩy ảnh nền chung.')
            data = self.read_body()
            if not data:
                raise ApiError(400, 'File rỗng.')
            src_sha = sha256_bytes(data)
            old = con.execute('SELECT src_sha FROM files WHERE kind=? AND path=?',
                              (kind, rel)).fetchone()
            if old and old['src_sha'] == src_sha and os.path.isfile(store_path(kind, rel)):
                return self.send_json({'ok': True, 'path': rel, 'unchanged': True})
            ext = os.path.splitext(rel)[1]
            if kind == 'avatars':
                data = shrink_avatar(data, ext)
            atomic_write(store_path(kind, rel), data)
            with _write_lock:
                record_file(con, kind, rel, data, src_sha, u['id'])
                con.commit()
            return self.send_json({'ok': True, 'path': rel, 'size': len(data),
                                   'sha': sha256_bytes(data), 'unchanged': False})
        if m == 'DELETE':
            if kind == 'avatars':
                self.need_admin(u)
            elif not self.can_write_bg(u, rel):
                raise ApiError(403, 'Không xoá được ảnh này.')
            try:
                os.remove(store_path(kind, rel))
            except FileNotFoundError:
                pass
            with _write_lock:
                con.execute('DELETE FROM files WHERE kind=? AND path=?', (kind, rel))
                audit(con, u['id'], f'delete-{kind}', rel)
                con.commit()
            return self.send_json({'ok': True})
        raise ApiError(405, 'Phương thức không hỗ trợ.')

    # ── FILE TỔNG ──
    @staticmethod
    def excel_for(con, uid):
        row = con.execute('SELECT * FROM excel WHERE owner=?', (str(uid),)).fetchone()
        return row or con.execute("SELECT * FROM excel WHERE owner='shared'").fetchone()

    @staticmethod
    def excel_public(row):
        return {'owner': row['owner'], 'file_name': row['file_name'], 'size': row['size'],
                'sha': row['sha'], 'updated_at': row['updated_at']}

    def api_excel_meta(self, con, u):
        row = self.excel_for(con, u['id'])
        if not row:
            raise ApiError(404, 'Chưa có FILE TỔNG trên máy chủ cho tài khoản này.')
        self.send_json({'ok': True, 'excel': self.excel_public(row)})

    def api_excel_file(self, con, u, q):
        owner = q.get('owner')
        if owner:
            self.need_admin(u)
            row = con.execute('SELECT * FROM excel WHERE owner=?', (owner,)).fetchone()
        else:
            row = self.excel_for(con, u['id'])
        if not row:
            raise ApiError(404, 'Chưa có FILE TỔNG.')
        self.send_file('excel', f'{row["owner"]}.xlsx', CTYPES['.xlsx'],
                       download_name=row['file_name'])

    def api_excel_put(self, con, u, q):
        target = (q.get('target') or 'me').strip()
        if target == 'me':
            owner = str(u['id'])
        elif target == 'all':
            self.need_admin(u)
            owner = 'shared'
        elif target.isdigit():
            self.need_admin(u)
            if not con.execute('SELECT 1 FROM users WHERE id=?', (int(target),)).fetchone():
                raise ApiError(404, 'Không có tài khoản này.')
            owner = target
        else:
            raise ApiError(400, 'target không hợp lệ.')
        name = urllib.parse.unquote(self.headers.get('X-File-Name') or '') or 'FILE_TONG.xlsx'
        name = os.path.basename(unicodedata.normalize('NFC', name).replace('\\', '/'))[:200]
        data = self.read_body()
        if not data.startswith(b'PK'):
            raise ApiError(400, 'File không phải Excel .xlsx.')
        atomic_write(os.path.join(STORE, 'excel', f'{owner}.xlsx'), data)
        sha = sha256_bytes(data)
        with _write_lock:
            con.execute(
                'INSERT INTO excel (owner, file_name, size, sha, updated_at, updated_by) VALUES (?,?,?,?,?,?) '
                'ON CONFLICT(owner) DO UPDATE SET file_name=excluded.file_name, size=excluded.size, '
                'sha=excluded.sha, updated_at=excluded.updated_at, updated_by=excluded.updated_by',
                (owner, name, len(data), sha, now_iso(), u['id']))
            cleared = 0
            if owner == 'shared':
                # «Đẩy cho TẤT CẢ»: bỏ các bản riêng để mọi tài khoản cùng nhận file chung.
                for r in con.execute("SELECT owner FROM excel WHERE owner<>'shared'").fetchall():
                    try:
                        os.remove(os.path.join(STORE, 'excel', f'{r["owner"]}.xlsx'))
                    except FileNotFoundError:
                        pass
                    cleared += 1
                con.execute("DELETE FROM excel WHERE owner<>'shared'")
            audit(con, u['id'], 'excel', f'{owner} {name} {len(data)}B')
            con.commit()
        row = con.execute('SELECT * FROM excel WHERE owner=?', (owner,)).fetchone()
        msg = ('Đã đẩy FILE TỔNG cho tất cả tài khoản.' if owner == 'shared'
               else 'Đã đẩy FILE TỔNG lên máy chủ.')
        self.send_json({'ok': True, 'msg': msg, 'excel': self.excel_public(row),
                        'cleared_personal': cleared})

    # ── quản trị ──
    def api_admin(self, con, u, p, q):
        m = self.command
        if p == '/users' and m == 'GET':
            rows = con.execute('SELECT * FROM users ORDER BY name COLLATE NOCASE, email').fetchall()
            ex = {r['owner'] for r in con.execute('SELECT owner FROM excel')}
            out = []
            for r in rows:
                d = user_public(r)
                d['personal_excel'] = str(r['id']) in ex
                out.append(d)
            return self.send_json({'ok': True, 'users': out})
        if p == '/users' and m == 'POST':
            return self.admin_save_user(con, u)
        mm = re.match(r'^/users/(\d+)/password$', p)
        if mm and m == 'POST':
            uid = int(mm.group(1))
            target = con.execute('SELECT * FROM users WHERE id=?', (uid,)).fetchone()
            if not target:
                raise ApiError(404, 'Không có tài khoản này.')
            if target['role'] == 'super_admin' and u['role'] != 'super_admin':
                raise ApiError(403, 'Không đổi được mật khẩu Super Admin.')
            pw = new_temp_password()
            with _write_lock:
                con.execute('UPDATE users SET pass_hash=?, must_change=1 WHERE id=?',
                            (hash_password(pw), uid))
                con.execute('DELETE FROM sessions WHERE user_id=?', (uid,))
                audit(con, u['id'], 'reset-password', target['email'])
                con.commit()
            return self.send_json({'ok': True, 'temp_password': pw})
        mm = re.match(r'^/users/(\d+)$', p)
        if mm and m == 'DELETE':
            uid = int(mm.group(1))
            if uid == u['id']:
                raise ApiError(400, 'Không tự xoá tài khoản của mình.')
            target = con.execute('SELECT * FROM users WHERE id=?', (uid,)).fetchone()
            if not target:
                raise ApiError(404, 'Không có tài khoản này.')
            if target['role'] == 'super_admin':
                raise ApiError(403, 'Không xoá được Super Admin.')
            with _write_lock:
                con.execute('DELETE FROM users WHERE id=?', (uid,))
                con.execute('DELETE FROM excel WHERE owner=?', (str(uid),))
                audit(con, u['id'], 'delete-user', target['email'])
                con.commit()
            try:
                os.remove(os.path.join(STORE, 'excel', f'{uid}.xlsx'))
            except FileNotFoundError:
                pass
            return self.send_json({'ok': True})
        if p == '/stats' and m == 'GET':
            out = {'ok': True}
            for kind in KINDS:
                r = con.execute('SELECT COUNT(*) n, COALESCE(SUM(size),0) b, MAX(updated_at) t '
                                'FROM files WHERE kind=?', (kind,)).fetchone()
                out[kind] = {'count': r['n'], 'bytes': r['b'], 'updated_at': r['t']}
            out['excel'] = [self.excel_public(r) for r in
                            con.execute("SELECT * FROM excel ORDER BY owner='shared' DESC, owner")]
            out['users'] = con.execute('SELECT COUNT(*) FROM users').fetchone()[0]
            out['audit'] = [dict(r) for r in con.execute(
                'SELECT a.ts, a.action, a.detail, u.email FROM audit a LEFT JOIN users u '
                'ON u.id=a.user_id ORDER BY a.rowid DESC LIMIT 40')]
            usage = shutil.disk_usage(DATA)
            out['disk'] = {'free': usage.free, 'total': usage.total}
            out['shrink'] = {'enabled': bool(AVATAR_MAX_PX and Image is not None),
                             'max_px': AVATAR_MAX_PX}
            return self.send_json(out)
        if p == '/report' and m == 'GET':
            owner = q.get('owner') or 'shared'
            row = con.execute('SELECT * FROM excel WHERE owner=?', (owner,)).fetchone()
            if not row:
                raise ApiError(404, 'Chưa có FILE TỔNG để đối chiếu.')
            res = build_report(os.path.join(STORE, 'excel', f'{owner}.xlsx'), row['sha'])
            res = dict(res)
            res['excel'] = self.excel_public(row)
            return self.send_json(res, status=200 if res.get('ok') else 500)
        if p == '/reindex' and m == 'POST':
            return self.send_json({'ok': True, **reindex()})
        raise ApiError(404, 'Không có API quản trị này.')

    def admin_save_user(self, con, u):
        d = self.read_json()
        uid = d.get('id')
        email = str(d.get('email') or '').strip().lower()
        name = str(d.get('name') or '').strip()[:120]
        role = str(d.get('role') or 'user')
        disabled = 1 if d.get('disabled') else 0
        if role not in ROLES:
            raise ApiError(400, 'Vai trò không hợp lệ.')
        if role == 'super_admin' and u['role'] != 'super_admin':
            raise ApiError(403, 'Chỉ Super Admin cấp được quyền Super Admin.')
        if not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', email):
            raise ApiError(400, 'Email không hợp lệ.')
        temp = None
        with _write_lock:
            if uid:
                target = con.execute('SELECT * FROM users WHERE id=?', (int(uid),)).fetchone()
                if not target:
                    raise ApiError(404, 'Không có tài khoản này.')
                if target['role'] == 'super_admin' and u['role'] != 'super_admin':
                    raise ApiError(403, 'Không sửa được Super Admin.')
                if target['id'] == u['id'] and (role != u['role'] or disabled):
                    raise ApiError(400, 'Không tự hạ quyền hoặc khoá chính mình.')
                try:
                    con.execute('UPDATE users SET email=?, name=?, role=?, disabled=? WHERE id=?',
                                (email, name, role, disabled, target['id']))
                except sqlite3.IntegrityError:
                    raise ApiError(400, 'Email đã có tài khoản khác dùng.')
                if disabled:
                    con.execute('DELETE FROM sessions WHERE user_id=?', (target['id'],))
                audit(con, u['id'], 'update-user', f'{email} {role}')
                new_id = target['id']
            else:
                temp = new_temp_password()
                try:
                    cur = con.execute(
                        'INSERT INTO users (email, name, role, pass_hash, must_change, created_at) '
                        'VALUES (?,?,?,?,1,?)', (email, name, role, hash_password(temp), now_iso()))
                except sqlite3.IntegrityError:
                    raise ApiError(400, 'Email này đã có tài khoản.')
                new_id = cur.lastrowid
                audit(con, u['id'], 'create-user', f'{email} {role}')
            con.commit()
        row = con.execute('SELECT * FROM users WHERE id=?', (new_id,)).fetchone()
        self.send_json({'ok': True, 'user': user_public(row), 'temp_password': temp})


# ═══════════════════════════ dòng lệnh ═══════════════════════════
def cli_create_admin(email, name=''):
    init_db()
    email = email.strip().lower()
    pw = new_temp_password()
    con = db()
    with _write_lock:
        row = con.execute('SELECT id FROM users WHERE email=?', (email,)).fetchone()
        if row:
            con.execute("UPDATE users SET role='super_admin', pass_hash=?, must_change=1, disabled=0 "
                        'WHERE id=?', (hash_password(pw), row['id']))
        else:
            con.execute('INSERT INTO users (email, name, role, pass_hash, must_change, created_at) '
                        "VALUES (?,?,'super_admin',?,1,?)",
                        (email, name or 'Quản trị', hash_password(pw), now_iso()))
        con.commit()
    con.close()
    print(f'Super Admin: {email}\nMật khẩu tạm (đổi ngay sau khi đăng nhập): {pw}')


def cli_set_password(email):
    init_db()
    pw = new_temp_password()
    con = db()
    with _write_lock:
        cur = con.execute('UPDATE users SET pass_hash=?, must_change=1 WHERE email=?',
                          (hash_password(pw), email.strip().lower()))
        con.execute('DELETE FROM sessions WHERE user_id IN (SELECT id FROM users WHERE email=?)',
                    (email.strip().lower(),))
        con.commit()
    con.close()
    if not cur.rowcount:
        print('Không có tài khoản này.')
        sys.exit(1)
    print(f'Mật khẩu tạm mới cho {email}: {pw}')


def main(argv):
    for stream in (sys.stdout, sys.stderr):      # console Windows không in được tiếng Việt
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    cmd = argv[1] if len(argv) > 1 else 'serve'
    if cmd == 'reindex':
        res = reindex()
        if len(argv) >= 3:
            res['src_sha_set'] = apply_src_map(argv[2])
        print(json.dumps(res, ensure_ascii=False))
        return
    if cmd == 'create-admin' and len(argv) >= 3:
        cli_create_admin(argv[2], ' '.join(argv[3:]))
        return
    if cmd == 'set-password' and len(argv) >= 3:
        cli_set_password(argv[2])
        return
    if cmd != 'serve':
        print(__doc__)
        sys.exit(2)
    init_db()
    srv = ThreadingHTTPServer((BIND, PORT), Handler)
    srv.daemon_threads = True
    print(f'AutoPPTX store: http://{BIND}:{PORT}  data={DATA}  '
          f'shrink={"on" if AVATAR_MAX_PX and Image else "off"}', flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    main(sys.argv)
