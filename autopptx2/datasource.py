"""Nguồn dữ liệu: Excel (openpyxl — không cần pandas), thư mục ảnh, avatar.

Không import Tkinter — mọi hàm gọi được từ worker thread.
"""
import os
import re
import difflib
from collections import OrderedDict

from .constants import IMG_EXTS, CODE_SUFFIX_RE


def _norm_header(s):
    """Chuẩn hoá tên cột: bỏ khoảng trắng/gạch/ngoặc/xuống dòng, thường hoá."""
    s = str(s or '').replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r"[\s_/\-\.,\(\)']+", '', s)
    return s.strip().lower()


# tên cột chuẩn → các biến thể chấp nhận (đã chuẩn hoá)
_COLUMN_ALIASES = {
    'Code_RP': ('coderp', 'codereport', 'reportcode', 'code', 'macode', 'mabaocao', 'marp',
                'rpcode', 'madiadiem', 'madiem', 'storecode', 'storeid',
                'siteid'),
    'Name': ('name', 'ten', 'tendiadiem', 'pointname', 'truong', 'nameofblock'),
    'Location': ('location', 'khuvuc', 'vitri'),
    'Address': ('address', 'diachi', 'addressmới', 'addressmoi'),
    'Ward': ('ward', 'phuong', 'wardmới', 'wardmoi'),
    'City': ('city', 'thanhpho', 'citymới', 'citymoi'),
    'District': ('district', 'quan', 'quanhuyen', 'distcũ', 'distcu'),
    'Channel': ('channel', 'kenh'),
    'Type': ('type', 'hinhthuc', 'format'),
    'Size': ('size', 'kichthuoc', 'inch', 'inches'),
    'GP_Size': ('gpsize', 'sizegp', 'kichthuocgp', 'giantpostersize'),
    'Note': ('note', 'notes', 'ghichu'),
    'Except': ('except', 'exceptbrand', 'exceptbrandcantadvertised'),
    'DP': ('dp', 'digitalposterinsideelevator'),
    'LCD': ('lcd', 'lcdfrontofelevator', 'lcdothers'),
    'GP': ('gp', 'giantposterinsideelevator', 'giantposteroutsidegroundfloor',
           'giantposteroutsideparkingfloor', 'giantposterstudyautox2forsale'),
    'GP_Inside': ('gpinside', 'qtyofgpinside', 'giantposterinsideelevator'),
    'GP_Ground': ('gpground', 'qtyofgpground', 'giantposteroutsidegroundfloor'),
    'GP_Facilities': ('gpfacilities', 'qtyofgpfacilitiesfloor'),
    'GP_Parking': ('gpparking', 'qtyofgpparkingfloor',
                   'giantposteroutsideparkingfloor'),
    'GP_Study': ('gpstudy', 'giantposterstudy', 'giantposterstudyautox2forsale'),
    'DS': ('ds', 'digitalposterothersdigitalstandee', 'digitalstandee'),
    'DPS': ('dps', 'digitalposterstair'),
    'DPF': ('dpf', 'digitalposterinfontoffloor', 'dpfdigitalposterinfontoffloor'),
    'LED': ('led',),
    'TrafficDay': ('trafficday', 'trafficngay'),
    'TrafficWeek': ('trafficweek', 'traffictuan', 'trafficwk'),
    'Quantity': ('quantity', 'soluong', 'qty'),
}

_QTY_KEYS = (
    'DP', 'LCD', 'GP', 'DS', 'DPS', 'DPF', 'LED',
    'GP_Inside', 'GP_Ground', 'GP_Facilities', 'GP_Parking', 'GP_Study',
)
_FORM_COLS = (
    ('DPS', 'DPS'),
    ('DP', 'DP'),
    ('DS', 'DS'),
    ('DPF', 'DPF'),
    ('LCD', 'LCD'),
    ('GP', 'GP'),
    ('LED', 'LED'),
)
_GP_FORM_COLS = (
    ('GP_Inside', 'GP in'),
    ('GP_Ground', 'GP đất'),
    ('GP_Facilities', 'GP tiện ích'),
    ('GP_Parking', 'GP gửi xe'),
    ('GP_Study', 'GP study'),
)
_FILL_KEYS = ('Code_RP', 'Name', 'Address', 'District', 'Channel',
              'Ward', 'City', 'TrafficDay', 'TrafficWeek', 'Type', 'Quantity')
_SUBHEADER_HINTS = ('led', 'digitalposter', 'giantposter', 'lcdfront', 'dpf',
                    'digitalstandee')


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


# Tỉnh/thành Bắc → Nam (Sales + BD). Khóa đã bỏ dấu.
_CITY_NS = (
    'ha giang', 'cao bang', 'bac kan', 'tuyen quang', 'lao cai', 'dien bien',
    'lai chau', 'son la', 'yen bai', 'hoa binh', 'thai nguyen', 'lang son',
    'quang ninh', 'bac giang', 'phu tho', 'vinh phuc', 'bac ninh', 'ha noi',
    'hai duong', 'hai phong', 'hung yen', 'thai binh', 'ha nam', 'nam dinh',
    'ninh binh',
    'thanh hoa', 'nghe an', 'ha tinh', 'quang binh', 'quang tri', 'hue',
    'da nang', 'quang nam', 'quang ngai', 'binh dinh', 'phu yen', 'khanh hoa',
    'gia lai', 'kon tum', 'dak lak', 'dak nong', 'lam dong',
    'ninh thuan', 'binh thuan', 'phan thiet',
    'binh phuoc', 'tay ninh', 'binh duong', 'dong nai', 'ba ria vung tau',
    'ho chi minh',
    'long an', 'tien giang', 'ben tre', 'tra vinh', 'vinh long', 'dong thap',
    'an giang', 'kien giang', 'can tho', 'hau giang', 'soc trang', 'bac lieu',
    'ca mau',
)
_CITY_INDEX = {name: i for i, name in enumerate(_CITY_NS)}
_CITY_STRIP = re.compile(
    r'^(thanh pho|tinh|tp\.?)\s+', re.IGNORECASE)
_VI_ASCII = str.maketrans({
    'à': 'a', 'á': 'a', 'ạ': 'a', 'ả': 'a', 'ã': 'a',
    'â': 'a', 'ầ': 'a', 'ấ': 'a', 'ậ': 'a', 'ẩ': 'a', 'ẫ': 'a',
    'ă': 'a', 'ằ': 'a', 'ắ': 'a', 'ặ': 'a', 'ẳ': 'a', 'ẵ': 'a',
    'è': 'e', 'é': 'e', 'ẹ': 'e', 'ẻ': 'e', 'ẽ': 'e',
    'ê': 'e', 'ề': 'e', 'ế': 'e', 'ệ': 'e', 'ể': 'e', 'ễ': 'e',
    'ì': 'i', 'í': 'i', 'ị': 'i', 'ỉ': 'i', 'ĩ': 'i',
    'ò': 'o', 'ó': 'o', 'ọ': 'o', 'ỏ': 'o', 'õ': 'o',
    'ô': 'o', 'ồ': 'o', 'ố': 'o', 'ộ': 'o', 'ổ': 'o', 'ỗ': 'o',
    'ơ': 'o', 'ờ': 'o', 'ớ': 'o', 'ợ': 'o', 'ở': 'o', 'ỡ': 'o',
    'ù': 'u', 'ú': 'u', 'ụ': 'u', 'ủ': 'u', 'ũ': 'u',
    'ư': 'u', 'ừ': 'u', 'ứ': 'u', 'ự': 'u', 'ử': 'u', 'ữ': 'u',
    'ỳ': 'y', 'ý': 'y', 'ỵ': 'y', 'ỷ': 'y', 'ỹ': 'y',
    'đ': 'd',
})
_CITY_ALIAS = {
    'hcm': 'ho chi minh',
    'tphcm': 'ho chi minh',
    'tp hcm': 'ho chi minh',
    'sai gon': 'ho chi minh',
    'hn': 'ha noi',
    'ha noi': 'ha noi',
    'thanh pho hue': 'hue',
    'ba ria - vung tau': 'ba ria vung tau',
    'brvt': 'ba ria vung tau',
    'vung tau': 'ba ria vung tau',
    'da lat': 'lam dong',
    'nha trang': 'khanh hoa',
}


def norm_city(city):
    """Tên tỉnh/thành để sắp xếp: bỏ dấu, bỏ tiền tố TP/Tỉnh."""
    s = str(clean(city) or '').strip().casefold().translate(_VI_ASCII)
    s = _CITY_STRIP.sub('', s)
    s = re.sub(r'[\s_\-./]+', ' ', s).strip()
    return _CITY_ALIAS.get(s, s)


def city_sort_key(city):
    """(có_tên, thứ_tự Bắc→Nam, tên). Thành phố trống xếp cuối."""
    key = norm_city(city)
    if not key:
        return (1, 999, '')
    return (0, _CITY_INDEX.get(key, 800), key)


def sort_codes_by_city(codes, by_code, merged=None):
    """Mã ảnh / Code_RP theo tỉnh-thành Bắc → Nam, rồi quận, tên."""
    codes = list(codes or [])
    by_code = by_code or {}
    if merged is None:
        merged = build_merged_groups(by_code)

    def key(code):
        row, _, _ = match_row(code, by_code, merged)
        row = row or {}
        return (
            city_sort_key(row.get('City')),
            str(clean(row.get('District')) or '').casefold(),
            str(clean(row.get('Name')) or '').casefold(),
            str(code or '').upper(),
        )

    return sorted(codes, key=key)


def order_groups_by_city(groups, by_code, merged=None):
    """OrderedDict nhóm ảnh theo tỉnh/thành (Sales + BD)."""
    groups = groups or {}
    codes = sort_codes_by_city(groups.keys(), by_code, merged)
    return OrderedDict((c, groups[c]) for c in codes if c in groups)


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
            self.rows = self._worksheet_rows(ws)
            wb.close()
            if not self.rows:
                self.error = "Không tìm thấy dòng tiêu đề chứa cột Code_RP."
                return False
            return True
        except Exception as e:
            self.error = str(e)
            return False

    def load_compare_list(self, path):
        """Đọc file đối chiếu; ưu tiên gộp mọi sheet có chữ ``List``.

        File nguồn University/Building thường tách ``University List`` và
        ``Building List``. Nếu file không có quy ước đó, dùng sheet đầu như
        luồng Excel thông thường để tránh kéo nhầm sheet Summary/AP OFF.
        """
        from openpyxl import load_workbook
        self.path, self.rows, self.error = path, [], None
        try:
            self.mtime = os.path.getmtime(path)
            wb = load_workbook(path, read_only=True, data_only=True)
            selected = [ws for ws in wb.worksheets if 'list' in ws.title.casefold()]
            if not selected and wb.worksheets:
                selected = [wb.worksheets[0]]
            for ws in selected:
                for rec in self._worksheet_rows(ws):
                    rec['_SourceSheet'] = ws.title
                    self.rows.append(rec)
            wb.close()
            if not self.rows:
                self.error = "Không tìm thấy dữ liệu Code_RP trong các sheet danh sách."
                return False
            return True
        except Exception as e:
            self.error = str(e)
            return False

    def _worksheet_rows(self, ws):
        """Chuẩn hoá các dòng dữ liệu từ một worksheet."""
        all_rows = list(ws.iter_rows(values_only=True))
        header_idx, header_map = None, None
        for i, row in enumerate(all_rows[:15]):
            cand = self._map_columns(row, require_code=True)
            if not cand:
                continue
            header_idx, header_map = i, cand
            if i + 1 < len(all_rows) and self._looks_like_subheader(all_rows[i + 1]):
                extra = self._map_columns(all_rows[i + 1], require_code=False)
                for idx, key in extra.items():
                    header_map.setdefault(idx, key)
                header_idx = i + 1
            break
        if header_map is None:
            return []
        rows, prev = [], {}
        for row in all_rows[header_idx + 1:]:
            rec = self._row_record(row, header_map)
            has_code = bool(str(rec.get('Code_RP') or '').strip())
            has_loc = bool(str(clean(rec.get('Location') or '')).strip())
            has_qty = any(to_qty(rec.get(k)) > 0 for k in _QTY_KEYS)
            if not (has_code or has_loc or has_qty):
                continue
            for k in _FILL_KEYS:
                if _is_blank(rec.get(k)) and not _is_blank(prev.get(k)):
                    rec[k] = prev[k]
            code = str(rec.get('Code_RP') or '').strip()
            if code and code.lower() != 'nan':
                rec['Code_RP'] = code
                rows.append(rec)
                prev = rec
        return rows

    @staticmethod
    def _map_columns(row, require_code=True):
        """{index: tên chuẩn}. require_code=True thì phải có Code_RP."""
        if not row:
            return {} if not require_code else None
        found = {}
        for idx, cell in enumerate(row):
            n = _norm_header(cell)
            if not n:
                continue
            for canon, aliases in _COLUMN_ALIASES.items():
                if n == canon.lower() or n in aliases:
                    found[idx] = canon
                    break
        if require_code:
            return found if 'Code_RP' in found.values() else None
        return found

    @staticmethod
    def _looks_like_subheader(row):
        texts = [_norm_header(c) for c in (row or []) if c]
        if not texts:
            return False
        if any(t in ('coderp', 'codereport') for t in texts):
            return False
        return any(any(h in t for h in _SUBHEADER_HINTS) for t in texts)

    @staticmethod
    def _row_record(row, header_map):
        rec = {}
        for idx, key in header_map.items():
            val = row[idx] if row and idx < len(row) else None
            if key in _QTY_KEYS:
                rec[key] = to_qty(rec.get(key)) + to_qty(val)
                continue
            if key in rec and not _is_blank(rec.get(key)):
                continue
            rec[key] = val
        return rec

    @staticmethod
    def _map_header(row):
        """Nếu `row` là dòng tiêu đề → {index: tên chuẩn}, ngược lại None."""
        return ExcelSource._map_columns(row, require_code=True)

    def channels(self):
        seen = []
        for r in self.rows:
            c = str(clean(r.get('Channel'))).strip()
            if c and c not in seen:
                seen.append(c)
        return sorted(seen)

    def by_code(self, channel=None):
        """{CODE (hoa): row đại diện} — lọc theo kênh nếu có.

        Nhiều dòng cùng mã (từng khu / loại màn) được cộng số lượng vào dòng đầu.
        """
        groups = OrderedDict()
        for r in self.rows:
            if channel:
                if str(clean(r.get('Channel'))).strip().casefold() != channel.casefold():
                    continue
            groups.setdefault(r['Code_RP'].upper(), []).append(r)
        out = {}
        for key, rs in groups.items():
            head = dict(rs[0])
            for qk in _QTY_KEYS:
                head[qk] = sum(to_qty(x.get(qk)) for x in rs)
            if _is_blank(head.get('Quantity')):
                for x in rs[1:]:
                    if not _is_blank(x.get('Quantity')):
                        head['Quantity'] = x.get('Quantity')
                        break
            out[key] = head
        return out

    def rows_of(self, code, channel=None):
        """Mọi dòng Excel cùng mã (từng khu vực / loại màn)."""
        key = str(code or '').strip().upper()
        if not key:
            return []
        out = []
        for r in self.rows:
            if str(r.get('Code_RP') or '').strip().upper() != key:
                continue
            if channel:
                if str(clean(r.get('Channel'))).strip().casefold() != channel.casefold():
                    continue
            out.append(r)
        return out


def inspect_excel(path):
    """Đọc nhanh file list chuẩn — không đụng UI. Trả dict kiểm tra."""
    src = ExcelSource()
    if not src.load(path):
        return {'ok': False, 'error': src.error or 'Không đọc được Excel.'}
    found = set()
    for r in src.rows:
        found.update(r.keys())
    recommended = ('Name', 'Address', 'District', 'Location', 'Channel',
                   'LCD', 'DP', 'DS', 'GP', 'DPS', 'TrafficDay', 'TrafficWeek')
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
# Hậu tố vị trí dán trên tên file — không thuộc Code_RP.
# GP*: decal GP từng mặt/tầng (GPI trong thang, GPG đất, GPF tiện ích/mặt thang,
# GPP gửi xe, GPS study). LCD/DP/DS/LED cùng kiểu đặt tên.
# Ví dụ CDCONGNGHETHUDUCGPG.jpg → CDCONGNGHETHUDUC.
_PLACE_SUFFIX_RE = re.compile(
    r'[\s_\-\.]?('
    r'GPSTUDY|GPINSIDE|GPGROUND|GPFACILITIES|GPPARKING|'
    r'DPS|DPF|LCD|LED|'
    r'GP[\s._-]?[SGPIF]|'
    r'GPS|GPG|GPF|GPI|GPP|GP|'
    r'DS|DP'
    r')\d*$',
    re.IGNORECASE,
)
_PLACE_BASE_MIN = 5


def strip_place_suffix(code):
    """Bỏ hậu tố vị trí (GPG/GPF/GPI…). Trả (mã gốc, hậu tố)."""
    raw = str(code or '').strip()
    if not raw:
        return '', ''
    m = _PLACE_SUFFIX_RE.search(raw)
    if not m:
        return raw, ''
    base = raw[:m.start()]
    if len(base) < _PLACE_BASE_MIN:
        return raw, ''
    return base, raw[m.start():]


def split_image_stem(stem):
    """'SCHOOLGPF (2)' → ('SCHOOL', 'GPF (2)')."""
    stem = str(stem or '')
    copy = ''
    m = CODE_SUFFIX_RE.search(stem)
    if m:
        copy = m.group(1)
        stem = stem[:m.start()]
    base, place = strip_place_suffix(stem)
    return base, f'{place}{copy}'


def image_group_code(path):
    """Mã nhóm ảnh từ tên file: bỏ (1)/_1 và hậu tố GP/LCD…"""
    stem = os.path.splitext(os.path.basename(str(path or '')))[0]
    base, _extra = split_image_stem(stem)
    return base


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
            code = image_group_code(path)
            if not code:
                continue
            g.setdefault(code, []).append(path)
        self._cache = (sig, g)
        return g

    def invalidate(self):
        """Bỏ cache nhóm — gọi sau khi đổi tên file trên đĩa."""
        self._cache = None

    def group_paths(self, code):
        return list(self.groups().get(code, []))

    @property
    def folder(self):
        return self.folders[0] if self.folders else None


# Hậu tố tên file avatar — quy ước V1: {Code_RP}, {Code_RP}A, {Code_RP}AA, {Code_RP}_Ava
_AVATAR_FILE_SUF = re.compile(r'(_Ava|AA|A)$', re.IGNORECASE)
# Alias mã block/tháp → ảnh avatar của mã gốc (giống V1, không phải fuzzy).
_AVATAR_BASE_RE = re.compile(
    r'(BLOCK|THAP|SANH|KHOI|FLAT|TOWER).*$', re.IGNORECASE)
# Mã gần giống (TOPAZHOME2BBLOCKB1 ↔ B2). Không đủ gần → None, không lấy file đầu folder.
_AVATAR_FUZZY_CUTOFF = 0.87


def _avatar_alnum(s):
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())


class AvatarIndex:
    """Tìm avatar theo Code_RP: khớp chính xác, chuẩn hoá (bỏ dấu/gạch),
    hậu tố tên file, alias BLOCK/THÁP, rồi gần giống (difflib 0.87).

    Không có file và không có tên gần giống → None. Không lấy ảnh đầu thư mục.
    """

    def __init__(self):
        self.folder = None
        self._exact = {}
        self._loose = {}
        self._aliases = {}

    def set_aliases(self, mapping):
        """Bảng alias tùy chọn: {mã_tìm: mã_file}. Không ghi đè khớp chính xác."""
        self._aliases = {}
        for src, dst in (mapping or {}).items():
            a, b = str(src or '').strip().upper(), str(dst or '').strip().upper()
            if a and b:
                self._aliases[a] = b

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
                    p = os.path.join(root, fn)
                    stem = os.path.splitext(fn)[0].strip().upper()
                    if not stem:
                        continue
                    self._put(stem, p)
                    stripped = _AVATAR_FILE_SUF.sub('', stem).strip()
                    if stripped and stripped != stem:
                        self._put(stripped, p)
        except Exception:
            pass

    def _put(self, stem, path):
        self._exact.setdefault(stem, path)
        loose = _avatar_alnum(stem)
        if loose:
            self._loose.setdefault(loose, path)

    def _lookup(self, key):
        if not key:
            return None
        p = self._exact.get(key)
        if p:
            return p
        return self._loose.get(_avatar_alnum(key))

    def find(self, code):
        if not code:
            return None
        key = str(code).strip().upper()
        if not key:
            return None
        seen = set()
        queue = [key]
        alias = self._aliases.get(key)
        if alias:
            queue.append(alias)
        base = _AVATAR_BASE_RE.sub('', key).strip()
        if base and base != key:
            queue.append(base)
            alias2 = self._aliases.get(base)
            if alias2:
                queue.append(alias2)
        for k in queue:
            if k in seen:
                continue
            seen.add(k)
            p = self._lookup(k)
            if p:
                return p
        cand = difflib.get_close_matches(
            key, list(self._exact), n=1, cutoff=_AVATAR_FUZZY_CUTOFF)
        if cand:
            return self._exact[cand[0]]
        return None

    def resolve(self, *codes):
        """Thử lần lượt các mã (ảnh, Code_RP Excel). Mã trống bỏ qua."""
        for c in codes:
            p = self.find(c)
            if p:
                return p
        return None


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
    qty_vals = [to_qty(r.get('Quantity')) for r in rows
                if not _is_blank(r.get('Quantity'))]
    if qty_vals:
        base['Quantity'] = sum(qty_vals)
    names = [str(r.get('Name', '') or '').strip() for r in rows if r.get('Name')]
    if names:
        common = os.path.commonprefix(names).rstrip(' -–—_')
        common = re.sub(r'\s+(Block|Tháp|Thap|Zone|Toà|Toa|Khu)\s*\w*$', '',
                        common, flags=re.IGNORECASE).strip()
        base['Name'] = common if len(common) >= 4 else names[0]
    return base


def build_photo_rename_plan(groups, mapping):
    """Lập kế hoạch đổi tên file theo {mã cũ: mã mới}.

    Giữ hậu tố vị trí (GPG/GPF…) + ' (2)' / '_2' và phần mở rộng.
    Không ghi đè file đã có. groups: {mã: [đường dẫn tuyệt đối]}.
    Trả (plan, conflicts) — mỗi phần tử (src, dst).
    """
    plan, conflicts = [], []
    planned = set()
    for old, new in (mapping or {}).items():
        old = str(old or '').strip()
        new = str(new or '').strip()
        if not old or not new:
            continue
        for src in groups.get(old, []) or []:
            folder = os.path.dirname(src)
            fn = os.path.basename(src)
            base, ext = os.path.splitext(fn)
            _code, extra = split_image_stem(base)
            new_fn = '{}{}{}'.format(new, extra, ext)
            if new_fn == fn:
                continue
            dst = os.path.join(folder, new_fn)
            key = os.path.normcase(os.path.normpath(dst))
            if os.path.exists(dst) or key in planned:
                conflicts.append((src, dst))
            else:
                plan.append((src, dst))
                planned.add(key)
    return plan, conflicts


def match_row(code, by_code, merged=None):
    """Khớp mã ảnh với dòng Excel.

    Trả (row|{}, mã_khớp|None, loại) — loại: exact | merged | fuzzy | none.
    """
    if not by_code:
        return {}, None, 'none'
    key = str(code or '').strip().upper()
    if not key:
        return {}, None, 'none'
    if key in by_code:
        return by_code[key], key, 'exact'
    base, sfx = strip_place_suffix(key)
    base = str(base or '').strip().upper()
    if sfx and base in by_code:
        return by_code[base], base, 'exact'
    if merged and key in merged:
        return merge_rows(merged[key]), key, 'merged'
    if merged and sfx and base in merged:
        return merge_rows(merged[base]), base, 'merged'
    cand = difflib.get_close_matches(key, list(by_code), n=1, cutoff=0.84)
    if not cand and sfx and base:
        cand = difflib.get_close_matches(base, list(by_code), n=1, cutoff=0.84)
    if cand:
        return by_code[cand[0]], cand[0], 'fuzzy'
    return {}, None, 'none'


# ════════════════════════════════════════════════════════════
#  Khối Thông tin — dựng chung cho preview & export
# ════════════════════════════════════════════════════════════
def screen_parts(row):
    """[(số, nhãn)] các loại màn > 0."""
    row = row or {}
    out = []
    gp_parts = []
    for key, lab in _GP_FORM_COLS:
        n = to_qty(row.get(key))
        if n > 0:
            gp_parts.append((n, lab))
    is_building = 'building' in str(clean(row.get('Channel'))).strip().casefold()
    # Building's generic LCD/GP fields are legacy aggregates (LCD is often
    # lifts and GP is the AP total), so placement-level GP is authoritative.
    if is_building and gp_parts:
        return gp_parts
    for key, lab in _FORM_COLS:
        n = to_qty(row.get(key))
        if n > 0:
            out.append((n, lab))
    if not out:
        out = gp_parts
    return out


def screen_qty(row):
    """Số màn tại điểm: LCD + DP + GP + DS (+ DPS/DPF/LED nếu có).

    Không dùng cột Quantity — list khách sạn thường ghi số phòng vào đó.
    """
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


# Thứ tự loại màn trong ngoặc Quantity — LCD trước DP như ví dụ V1
# "8 (2 LCD, 6 DP)"; DS/GP sau; DPS/DPF/LED nếu có.
_QTY_SHOW_ORDER = ('LCD', 'DP', 'DS', 'GP', 'DPS', 'DPF', 'LED')


def _combined_screen_parts(row):
    """Return non-GP screen forms in the same order as Sales Quantity."""
    row = row or {}
    rank = {label: index for index, label in enumerate(_QTY_SHOW_ORDER)}
    parts = []
    for key, label in _FORM_COLS:
        if key == 'GP':
            continue
        qty = to_qty(row.get(key))
        if qty > 0:
            parts.append((label, qty))
    return sorted(parts, key=lambda item: rank.get(item[0], 99))


def _combined_screen_summary(parts):
    """Format DP/LCD-family forms as one Quantity-style table value."""
    if not parts:
        return ''
    total = sum(qty for _, qty in parts)
    detail = ', '.join(f'{qty} {label}' for label, qty in parts)
    return f'{total} ({detail})'


def _quantity_text(row, n_files=0):
    """Quantity trên slide = tổng màn + tách loại, giống AutoPPTX cũ.

    V1 ghép `{tổng} ({n DP}, {n LCD}, …)` từ cột LCD/DP/GP/DS — không lấy
    nguyên chuỗi từ cột Quantity. Cột Quantity Excel chỉ dùng khi mọi cột
    loại màn = 0 (Coffee/Beauty). Không lấy số phòng khi đã có số màn.
    """
    row = row or {}
    parts = screen_parts(row)
    n = sum(c for c, _ in parts)
    if n <= 0:
        n = to_qty(row.get('so_man_slot'))
    if n > 0:
        if parts:
            rank = {lab: i for i, lab in enumerate(_QTY_SHOW_ORDER)}
            parts = sorted(parts, key=lambda p: rank.get(p[1], 99))
            return '{} ({})'.format(
                n, ', '.join('{} {}'.format(c, lab) for c, lab in parts))
        return str(n)
    raw = row.get('Quantity')
    if not _is_blank(raw) and to_qty(raw) > 0:
        return fmt_num(raw)
    try:
        nf = int(n_files or 0)
    except (TypeError, ValueError):
        nf = 0
    if nf > 0:
        return str(nf)
    return '—'


def build_info_rows(row, n_files=0):
    """[(nhãn, giá trị, là_dòng_nhấn)] cho khối Thông tin mẫu báo cáo cũ."""
    row = row or {}
    return [
        ("Address", str(clean(row.get('Address'))), False),
        ("District", str(clean(row.get('District'))), False),
        ("Quantity", _quantity_text(row, n_files), False),
        ("Traffic / day", fmt_num(row.get('TrafficDay')), True),
        ("Traffic / week", fmt_num(row.get('TrafficWeek')), True),
    ]


def _note_text(row):
    """Note trên bảng: ưu tiên hạn chế (Except). Ẩn trạng thái ON AIR."""
    row = row or {}
    exc = str(clean(row.get('Except'))).strip()
    if exc:
        return exc
    note = str(clean(row.get('Note'))).strip()
    if note.upper().replace('-', ' ') in ('ON AIR', 'OFF AIR', 'ONGOING', 'ON GOING'):
        return ''
    return note


def _khu_vuc(row):
    loc = str(clean((row or {}).get('Location'))).strip()
    if ' - ' in loc:
        return loc.split(' - ', 1)[1].strip()
    if loc:
        return loc
    return str(clean((row or {}).get('District'))).strip()


def _school_short(row):
    loc = str(clean((row or {}).get('Location'))).strip()
    if ' - ' in loc:
        return loc.split(' - ', 1)[0].strip()
    return str(clean((row or {}).get('Name'))).strip()


def _address_line(row):
    row = row or {}
    addr = str(clean(row.get('Address'))).strip()
    ward = str(clean(row.get('Ward') or row.get('District'))).strip()
    city = str(clean(row.get('City'))).strip()
    parts = []
    if addr:
        parts.append(addr)
    if ward:
        wl = ward.lower()
        if wl[:1].isdigit():
            parts.append('Quận ' + ward)
        elif wl.startswith(('phường', 'phuong', 'quận', 'quan', 'p.', 'q.', 'tp')):
            parts.append(ward)
        else:
            parts.append('Phường ' + ward)
    if city:
        cl = city.lower()
        if 'hồ chí minh' in cl or cl in ('hcm', 'tphcm', 'tp.hcm'):
            parts.append('TP.HCM')
        else:
            parts.append(city)
    return ', '.join(parts)


def _fmt_traffic(row):
    v = (row or {}).get('TrafficWeek')
    if _is_blank(v):
        v = (row or {}).get('TrafficDay')
    v = clean(v)
    if isinstance(v, (int, float)) and float(v) >= 100:
        return f'{int(round(float(v))):,}'.replace(',', '.')
    if v == '':
        return ''
    return fmt_num(v)


def build_info_table(row, siblings=None):
    """Khung bảng SALESKIT: ĐỊA ĐIỂM / ĐỊA CHỈ / TRAFFIC + từng khu-loại màn."""
    siblings = [r for r in (siblings or []) if r]
    if not siblings and row:
        siblings = [row]
    head = siblings[0] if siblings else (row or {})
    specs = []
    for r in siblings:
        area = _khu_vuc(r)
        screen_size = str(clean(r.get('Size'))).strip()
        gp_size = str(clean(r.get('GP_Size'))).strip()
        note = _note_text(r)
        detailed_gp = sum(to_qty(r.get(key)) for key, _ in _GP_FORM_COLS)
        gp_qty = to_qty(r.get('GP')) or detailed_gp
        screen_forms = _combined_screen_parts(r)
        if screen_forms:
            # Match the Sales Quantity presentation: all digital formats use
            # one row and share the screen Size from the master.
            specs.append({
                'area': area,
                'form': ', '.join(label for label, _ in screen_forms),
                'size': screen_size,
                'qty': _combined_screen_summary(screen_forms),
                'note': note,
            })
        if gp_qty > 0:
            # All GP placements are intentionally presented as one sale type.
            specs.append({
                'area': area, 'form': 'GP', 'size': gp_size,
                'qty': str(gp_qty), 'note': note,
            })
        if screen_forms or gp_qty > 0:
            continue
        qn = to_qty(r.get('Quantity'))
        if area or screen_size or note or qn > 0:
            fallback_form = str(clean(r.get('Type'))).strip()
            specs.append({
                'area': area,
                'form': fallback_form,
                'size': gp_size if fallback_form.upper() == 'GP' else screen_size,
                'qty': str(qn) if qn > 0 else '',
                'note': note,
            })
    if not specs:
        specs.append({'area': '', 'form': '', 'size': '', 'qty': '', 'note': ''})
    school = _school_short(head) or str(clean(head.get('Name'))).strip()
    return {
        'school': school,
        'address': _address_line(head),
        'traffic': _fmt_traffic(head),
        'specs': specs,
    }


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


_QUANTITY_IS_SCREEN_CHANNELS = (
    'coffee', 'cafe', 'cà phê', 'milk tea', 'milktea', 'trà sữa',
    'fastfood', 'fast food', 'beauty', 'salon', '30shine',
)


def expected_photo_count(row):
    """Số ảnh tối thiểu cần có cho một điểm trong báo cáo đối chiếu.

    Ưu tiên tổng số màn chuẩn. Với các kênh cửa hàng nơi ``Quantity`` là số
    màn, dùng ``Quantity`` nếu file không tách DP/LCD/GP. Các kênh còn lại vẫn
    yêu cầu ít nhất một ảnh để không bỏ sót điểm chỉ có thông tin định tính.
    """
    row = row or {}
    n = screen_qty(row)
    if n > 0:
        return n
    channel = str(clean(row.get('Channel'))).strip().casefold()
    if any(token in channel for token in _QUANTITY_IS_SCREEN_CHANNELS):
        n = to_qty(row.get('Quantity'))
        if n > 0:
            return n
    return 1


def compare_list_with_master(list_rows, master_by_code, image_groups, merged=None,
                             check_images=True):
    """Đối chiếu một file list với file tổng và ảnh, nhóm kết quả theo kênh.

    ``list_rows`` chỉ xác định các mã cần kiểm tra. ``master_by_code`` là file
    tổng đang mở/đã đồng bộ và cung cấp Channel, Name, số màn. ``image_groups``
    là kết quả quét thư mục ảnh hiện tại. Khi ``check_images`` là ``False``,
    chỉ đối chiếu Code_RP và không kết luận thiếu ảnh.
    """
    master_by_code = {
        str(code or '').strip().upper(): row
        for code, row in (master_by_code or {}).items()
        if str(code or '').strip()
    }
    merged = merged or {}
    unique = OrderedDict()
    for rec in list_rows or []:
        code = str((rec or {}).get('Code_RP') or '').strip()
        if code:
            unique.setdefault(code.upper(), rec or {})

    direct_photos = {}
    master_photos = {}
    for image_code, paths in (image_groups or {}).items():
        key = str(image_code or '').strip().upper()
        count = len(paths or [])
        if key:
            direct_photos[key] = direct_photos.get(key, 0) + count
        _row, matched, _kind = match_row(key, master_by_code, merged)
        if matched:
            mk = str(matched).strip().upper()
            master_photos[mk] = master_photos.get(mk, 0) + count

    rows = []
    channel_map = {}
    for code, list_row in unique.items():
        master_row, matched, kind = match_row(code, master_by_code, merged)
        has_master = bool(master_row) and kind != 'none'
        source = master_row if has_master else list_row
        canonical = str(matched or code).strip().upper()
        if check_images:
            photos = max(direct_photos.get(code, 0), master_photos.get(canonical, 0))
            required = expected_photo_count(source)
            missing = max(required - photos, 0)
        else:
            photos = required = missing = 0
        channel = str(clean((source or {}).get('Channel'))).strip() or 'Không xác định'
        name = str(clean((source or {}).get('Name'))).strip()
        district = str(clean((source or {}).get('District'))).strip()
        status = 'not_in_master' if not has_master else ('missing' if missing else 'ok')
        item = {
            'code': code,
            'matched_code': canonical if has_master else '',
            'name': name,
            'district': district,
            'channel': channel,
            'required': required,
            'photos': photos,
            'missing': missing,
            'status': status,
            'match_kind': kind,
        }
        rows.append(item)
        summary = channel_map.setdefault(channel, {
            'channel': channel, 'sites': 0, 'ok_sites': 0,
            'missing_sites': 0, 'not_in_master': 0,
            'required': 0, 'photos': 0, 'missing': 0,
        })
        summary['sites'] += 1
        summary['required'] += required
        summary['photos'] += photos
        summary['missing'] += missing
        if status == 'ok':
            summary['ok_sites'] += 1
        else:
            if missing:
                summary['missing_sites'] += 1
            if status == 'not_in_master':
                summary['not_in_master'] += 1

    order = {'missing': 0, 'not_in_master': 1, 'ok': 2}
    rows.sort(key=lambda r: (
        r['channel'].casefold(), order.get(r['status'], 9), r['code']))
    channels = sorted(channel_map.values(), key=lambda r: r['channel'].casefold())
    return {
        'images_checked': bool(check_images),
        'rows': rows,
        'channels': channels,
        'n_list': len(rows),
        'n_master_matches': sum(1 for r in rows if r['status'] != 'not_in_master'),
        'n_not_in_master': sum(1 for r in rows if r['status'] == 'not_in_master'),
        'n_ok_sites': sum(1 for r in rows if r['status'] == 'ok'),
        'n_missing_sites': sum(1 for r in rows if r['missing'] > 0),
        'n_required': sum(r['required'] for r in rows),
        'n_photos': sum(r['photos'] for r in rows),
        'n_missing_images': sum(r['missing'] for r in rows),
    }
