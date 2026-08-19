"""Nguồn dữ liệu: Excel (openpyxl — không cần pandas), thư mục ảnh, avatar.

Không import Tkinter — mọi hàm gọi được từ worker thread.
"""
import os
import re
import difflib
from collections import OrderedDict

from .constants import IMG_EXTS, CODE_SUFFIX_RE


def _norm_header(s):
    """Chuẩn hoá tên cột: bỏ khoảng trắng/gạch/chéo, thường hoá."""
    return re.sub(r'[\s_/\-\.]+', '', str(s or '')).strip().lower()


# tên cột chuẩn → các biến thể chấp nhận (đã chuẩn hoá)
_COLUMN_ALIASES = {
    'Code_RP': ('coderp', 'code', 'macode', 'mabaocao', 'marp', 'rpcode',
                'madiadiem', 'madiem', 'storecode', 'storeid', 'siteid'),
    'Name': ('name', 'ten', 'tendiadiem', 'pointname'),
    'Address': ('address', 'diachi'),
    'District': ('district', 'quan', 'quanhuyen'),
    'Channel': ('channel', 'kenh'),
    'DP': ('dp',),
    'LCD': ('lcd',),
    'GP': ('gp',),
    'DS': ('ds',),
    'TrafficDay': ('trafficday', 'trafficngay'),
    'TrafficWeek': ('trafficweek', 'traffictuan'),
}


def _is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and v != v:      # NaN
        return True
    return str(v).strip() == ''


def clean(v):
    """None/NaN → chuỗi rỗng; số nguyên dạng float → int."""
    if _is_blank(v):
        return ''
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def fmt_num(v):
    """Hiển thị số có phân tách hàng nghìn; chuỗi giữ nguyên."""
    v = clean(v)
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        return f"{v:,.1f}"
    return str(v)


def to_qty(v):
    if _is_blank(v):
        return 0
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return 0


class ExcelSource:
    """Đọc file .xlsx/.xlsm bằng openpyxl (nhanh, không kéo theo pandas)."""

    def __init__(self):
        self.path = None
        self.mtime = None
        self.rows = []           # list[dict] khoá chuẩn Code_RP, Name, ...
        self.error = None

    def load(self, path):
        """Đọc file — gọi từ worker thread. Trả True nếu thành công."""
        from openpyxl import load_workbook
        self.path, self.rows, self.error = path, [], None
        try:
            self.mtime = os.path.getmtime(path)
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            header_map = None       # index cột → tên chuẩn
            for row in ws.iter_rows(max_row=15, values_only=True):
                cand = self._map_header(row)
                if cand:
                    header_map = cand
                    break
            if header_map is None:
                self.error = "Không tìm thấy dòng tiêu đề chứa cột Code_RP."
                wb.close()
                return False
            started = False
            for row in ws.iter_rows(values_only=True):
                if not started:
                    # bỏ qua tới hết dòng tiêu đề
                    if self._map_header(row):
                        started = True
                    continue
                rec = {}
                for idx, key in header_map.items():
                    rec[key] = row[idx] if idx < len(row) else None
                code = str(rec.get('Code_RP') or '').strip()
                if code and code.lower() != 'nan':
                    rec['Code_RP'] = code
                    self.rows.append(rec)
            wb.close()
            return True
        except Exception as e:
            self.error = str(e)
            return False

    @staticmethod
    def _map_header(row):
        """Nếu `row` là dòng tiêu đề → {index: tên chuẩn}, ngược lại None."""
        if not row:
            return None
        found = {}
        for idx, cell in enumerate(row):
            n = _norm_header(cell)
            if not n:
                continue
            for canon, aliases in _COLUMN_ALIASES.items():
                if n in aliases and canon not in found.values():
                    found[idx] = canon
        return found if 'Code_RP' in found.values() else None

    def channels(self):
        seen = []
        for r in self.rows:
            c = str(clean(r.get('Channel'))).strip()
            if c and c not in seen:
                seen.append(c)
        return sorted(seen)

    def by_code(self, channel=None):
        """{CODE (hoa): row} — lọc theo kênh nếu có."""
        out = {}
        for r in self.rows:
            if channel:
                if str(clean(r.get('Channel'))).strip().casefold() != channel.casefold():
                    continue
            out[r['Code_RP'].upper()] = r
        return out


def inspect_excel(path):
    """Đọc nhanh file list chuẩn — không đụng UI. Trả dict kiểm tra."""
    src = ExcelSource()
    if not src.load(path):
        return {'ok': False, 'error': src.error or 'Không đọc được Excel.'}
    found = set()
    for r in src.rows:
        found.update(r.keys())
    recommended = ('Name', 'Address', 'District', 'Channel',
                   'LCD', 'DP', 'DS', 'GP', 'TrafficDay', 'TrafficWeek')
    missing = [k for k in recommended if k not in found]
    sample = []
    for r in src.rows[:8]:
        sample.append(str(r.get('Code_RP') or ''))
    return {
        'ok': True,
        'n': len(src.rows),
        'channels': src.channels(),
        'found': [k for k in ('Code_RP',) + recommended if k in found],
        'missing': missing,
        'sample_codes': sample,
        'error': '',
    }


# ════════════════════════════════════════════════════════════
#  Ảnh & avatar
# ════════════════════════════════════════════════════════════
class ImageLibrary:
    """Quét một hoặc nhiều thư mục ảnh, nhóm theo mã. Đường dẫn tuyệt đối."""

    def __init__(self):
        self.folders = []
        self.recursive = True
        self._cache = None          # (sig, groups)

    def set_folder(self, folder):
        """Tương thích cũ: một thư mục."""
        self.set_folders([folder] if folder else [], recursive=self.recursive)

    def set_folders(self, folders, recursive=True):
        seen, out = set(), []
        for f in folders or []:
            if not f:
                continue
            p = os.path.normpath(f)
            key = os.path.normcase(p)
            if key in seen or not os.path.isdir(p):
                continue
            seen.add(key)
            out.append(p)
        self.folders = out
        self.recursive = bool(recursive)
        self._cache = None

    def add_folder(self, folder, recursive=None):
        if recursive is not None:
            self.recursive = bool(recursive)
        folders = list(self.folders)
        folders.append(folder)
        self.set_folders(folders, recursive=self.recursive)

    def remove_folder(self, folder):
        key = os.path.normcase(os.path.normpath(folder or ''))
        self.set_folders(
            [f for f in self.folders if os.path.normcase(f) != key],
            recursive=self.recursive)

    def _sig(self):
        parts = []
        for f in self.folders:
            try:
                parts.append((f, os.path.getmtime(f), self.recursive))
            except Exception:
                parts.append((f, None, self.recursive))
        return tuple(parts)

    def _iter_images(self):
        seen = set()
        for folder in self.folders:
            try:
                if self.recursive:
                    walker = ((r, files) for r, _d, files in os.walk(folder))
                else:
                    walker = ((folder, os.listdir(folder)),)
                for root, files in walker:
                    for fn in files:
                        if not fn.lower().endswith(IMG_EXTS):
                            continue
                        path = os.path.join(root, fn)
                        key = os.path.normcase(os.path.normpath(path))
                        if key in seen:
                            continue
                        seen.add(key)
                        yield path
            except Exception:
                continue

    def groups(self):
        """OrderedDict {mã: [đường dẫn tuyệt đối]} theo thứ tự tên."""
        sig = self._sig()
        if self._cache and self._cache[0] == sig:
            return self._cache[1]
        g = OrderedDict()
        for path in sorted(self._iter_images(), key=lambda p: os.path.basename(p).casefold()):
            code = CODE_SUFFIX_RE.sub('', os.path.splitext(os.path.basename(path))[0])
            g.setdefault(code, []).append(path)
        self._cache = (sig, g)
        return g

    def group_paths(self, code):
        return list(self.groups().get(code, []))

    @property
    def folder(self):
        return self.folders[0] if self.folders else None


class AvatarIndex:
    """Tìm avatar theo mã điểm (khớp chính xác → chuẩn hoá → gần đúng)."""

    def __init__(self):
        self.folder = None
        self._exact = {}
        self._loose = {}

    def set_folder(self, folder):
        self.folder = folder
        self._exact, self._loose = {}, {}
        if not folder or not os.path.isdir(folder):
            return
        try:
            for root, _dirs, files in os.walk(folder):
                for fn in files:
                    if not fn.lower().endswith(IMG_EXTS):
                        continue
                    stem = os.path.splitext(fn)[0].strip()
                    stem = re.sub(r'(_Ava|AA|A)$', '', stem, flags=re.IGNORECASE).upper()
                    p = os.path.join(root, fn)
                    self._exact.setdefault(stem, p)
                    self._loose.setdefault(re.sub(r'[^A-Z0-9]', '', stem), p)
        except Exception:
            pass

    def find(self, code):
        if not code:
            return None
        key = str(code).strip().upper()
        p = self._exact.get(key)
        if p:
            return p
        p = self._loose.get(re.sub(r'[^A-Z0-9]', '', key))
        if p:
            return p
        cand = difflib.get_close_matches(key, list(self._exact), n=1, cutoff=0.87)
        return self._exact[cand[0]] if cand else None

    def first(self):
        return next(iter(self._exact.values()), None)


# ════════════════════════════════════════════════════════════
#  Khớp mã ảnh ↔ Excel
# ════════════════════════════════════════════════════════════
# Kênh Building Premium đặt tên ảnh RÚT GỌN (AKARICITY) trong khi Excel tách
# block (AKARICITYBLOCKAK1..). Suy ngược từ Code_RP để nối hai bên.
_BLOCK_SUFFIX = re.compile(
    r'^(.*?)(BLOCK.*|THAP\d+.*|VP[A-Z]?\d*|SKY\d+|SF\d+|S[AO]\d+|AK\d+|CS\d+)$')
_MERGE_CHANNELS = {'building premium'}
_SUM_COLS = ('LCD', 'DP', 'DS', 'GP', 'TrafficDay', 'TrafficWeek')


def build_merged_groups(by_code):
    """{MÃ_RÚT_GỌN: [rows]} cho kênh cần gộp block."""
    merged = {}
    for k, row in by_code.items():
        ch = str(clean(row.get('Channel'))).strip().casefold()
        if ch not in _MERGE_CHANNELS:
            continue
        m = _BLOCK_SUFFIX.match(k)
        root = m.group(1) if m else ''
        if len(root) >= 5 and root != k:
            merged.setdefault(root, []).append(row)
    for k in list(merged):
        if k in by_code:            # mã thật luôn thắng
            del merged[k]
    return merged


def merge_rows(rows):
    """Gộp nhiều block thành MỘT dòng: cộng số liệu, lấy phần chung của tên."""
    codes = [str(r.get('Code_RP', '') or '').strip().upper() for r in rows]
    base = dict(rows[0])
    base['_merged_codes'] = codes
    if len(rows) == 1:
        return base
    for c in _SUM_COLS:
        total = 0
        for r in rows:
            try:
                v = r.get(c)
                total += 0 if _is_blank(v) else float(v)
            except Exception:
                pass
        base[c] = int(total)
    names = [str(r.get('Name', '') or '').strip() for r in rows if r.get('Name')]
    if names:
        common = os.path.commonprefix(names).rstrip(' -–—_')
        common = re.sub(r'\s+(Block|Tháp|Thap|Zone|Toà|Toa|Khu)\s*\w*$', '',
                        common, flags=re.IGNORECASE).strip()
        base['Name'] = common if len(common) >= 4 else names[0]
    return base


def match_row(code, by_code, merged=None):
    """Khớp mã ảnh với dòng Excel.

    Trả (row|{}, mã_khớp|None, loại) — loại: exact | merged | fuzzy | none.
    """
    if not by_code:
        return {}, None, 'none'
    key = str(code or '').strip().upper()
    if key in by_code:
        return by_code[key], key, 'exact'
    if merged and key in merged:
        return merge_rows(merged[key]), key, 'merged'
    cand = difflib.get_close_matches(key, list(by_code), n=1, cutoff=0.84)
    if cand:
        return by_code[cand[0]], cand[0], 'fuzzy'
    return {}, None, 'none'


# ════════════════════════════════════════════════════════════
#  Khối Thông tin — dựng chung cho preview & export
# ════════════════════════════════════════════════════════════
def screen_parts(row):
    """[(số, nhãn)] các loại màn > 0 — LCD, DP, DS, GP."""
    row = row or {}
    out = []
    for key, lab in (('LCD', 'LCD'), ('DP', 'DP'), ('DS', 'DS'), ('GP', 'GP')):
        n = to_qty(row.get(key))
        if n > 0:
            out.append((n, lab))
    if not out:
        for key, lab in (('GP_Inside', 'GP in'), ('GP_Ground', 'GP đất'),
                         ('GP_Facilities', 'GP tiện ích'),
                         ('GP_Parking', 'GP gửi xe')):
            n = to_qty(row.get(key))
            if n > 0:
                out.append((n, lab))
    return out


def screen_qty(row):
    """Số tivi / màn hình kỳ vọng của 1 điểm (LCD + DP + DS + GP)."""
    parts = screen_parts(row)
    total = sum(n for n, _ in parts)
    if total <= 0:
        total = to_qty((row or {}).get('so_man_slot'))
    return total


def place_label(row):
    """Tên địa điểm đóng dấu / tiêu đề phụ: Name · District (từ Excel theo mã)."""
    row = row or {}
    name = str(clean(row.get('Name'))).strip()
    dist = str(clean(row.get('District'))).strip()
    addr = str(clean(row.get('Address'))).strip()
    if name and dist:
        return f'{name}\n{dist}'
    return name or addr or ''


def build_info_rows(row, n_files):
    """[(nhãn, giá trị, là_dòng_nhấn)] cho khối Thông tin.

    Photos = số ảnh đã chụp. Tivi = số màn trên list Excel (không đếm trong ảnh).
    """
    row = row or {}
    n_files = int(n_files or 0)
    tvs = screen_qty(row)
    parts = screen_parts(row)
    detail = ', '.join(f'{n} {lab}' for n, lab in parts)
    if tvs > 0:
        tv_txt = f'{tvs}' + (f'  ({detail})' if detail else '')
    else:
        tv_txt = '—'
    return [
        ("Address", str(clean(row.get('Address'))), False),
        ("District", str(clean(row.get('District'))), False),
        ("Photos", str(n_files), False),
        ("Tivi / màn", tv_txt, False),
        ("Traffic / day", fmt_num(row.get('TrafficDay')), True),
        ("Traffic / week", fmt_num(row.get('TrafficWeek')), True),
    ]


def channel_text(row, template):
    """Chuỗi phần tử Kênh từ template 'Report {channel} 2026'."""
    ch = str(clean((row or {}).get('Channel'))).strip() or "Channel"
    out = template or "Report {channel} 2026"
    for ph in ('{channel}', '{Channel}', '{CHANNEL}'):
        out = out.replace(ph, ch)
    return out


def _ov_status(photos, screens, kind):
    if kind == 'none' and screens <= 0:
        return 'chua_excel', 'Chưa có Excel'
    if photos <= 0:
        return 'thieu', 'THIẾU ảnh'
    if screens <= 0:
        return 'chua_man', 'Chưa có số màn'
    if photos < screens:
        return 'thieu', 'THIẾU ảnh'
    if photos > screens:
        return 'thua', 'Thừa ảnh'
    return 'du', 'ĐỦ'


def build_overview(groups, by_code, merged=None):
    """Đối chiếu từng cửa hàng: số ảnh đã up vs số màn hình trên Excel.

    Trả dict rows + tổng: du / thieu / thua / chua_excel / excel_chua_anh.
    """
    groups = groups or {}
    by_code = by_code or {}
    rows = []
    matched_excel = set()
    for code, paths in groups.items():
        row, mcode, kind = match_row(code, by_code, merged)
        photos = len(paths or [])
        screens = screen_qty(row)
        name = str(clean((row or {}).get('Name'))).strip() or str(code)
        district = str(clean((row or {}).get('District'))).strip()
        st, label = _ov_status(photos, screens, kind)
        if mcode:
            matched_excel.add(str(mcode).strip().upper())
        rows.append({
            'code': str(code),
            'name': name,
            'district': district,
            'photos': photos,
            'screens': screens,
            'delta': photos - screens,
            'status': st,
            'label': label,
            'kind': kind,
        })
    order = {'thieu': 0, 'chua_excel': 1, 'chua_man': 2, 'thua': 3, 'du': 4}
    rows.sort(key=lambda r: (order.get(r['status'], 9), r['code'].upper()))
    excel_no_photo = []
    if by_code:
        photo_keys = {str(c).strip().upper() for c in groups}
        for k, row in by_code.items():
            ku = str(k).strip().upper()
            if ku in matched_excel or ku in photo_keys:
                continue
            excel_no_photo.append(ku)
    excel_no_photo.sort()
    return {
        'rows': rows,
        'n_stores': len(rows),
        'n_photos': sum(r['photos'] for r in rows),
        'n_screens': sum(r['screens'] for r in rows),
        'n_du': sum(1 for r in rows if r['status'] == 'du'),
        'n_thieu': sum(1 for r in rows if r['status'] == 'thieu'),
        'n_thua': sum(1 for r in rows if r['status'] == 'thua'),
        'n_chua_excel': sum(1 for r in rows if r['status'] == 'chua_excel'),
        'excel_no_photo': excel_no_photo,
    }
