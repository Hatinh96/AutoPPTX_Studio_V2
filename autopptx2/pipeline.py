"""Dựng ExportJob không cần giao diện — cho máy chủ xuất PPTX (bản web).

Đi đúng các bước app desktop làm trước khi xuất (app.py: _rebuild_code_map,
_rebuild_scope_map, _filtered_groups, _ordered_groups, _resolve_avatar,
_start_export), để file xuất trên web giống hệt file xuất từ app.

Không import Tkinter / customtkinter.
"""
import copy
import os
from collections import OrderedDict

from .constants import (DEFAULT_LAYOUT, ELEMENT_NAMES, AR_CHOICES, ALL_CHANNELS,
                        ALL_CITIES, ALL_DISTRICTS, builtin_pack, slide_inches)
from .datasource import (ExcelSource, ImageLibrary, AvatarIndex, match_row,
                         build_merged_groups, scope_by_channel, filter_groups_by_geo,
                         order_groups, geo_field_values)
from .effects import DEFAULT_FX
from . import exporter
from . import geometry as G
from . import qa as QA

PRESET_NAMES = ('Báo cáo', 'SALESKIT', 'THE NELSON')


def deep_merge(base, extra):
    out = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def use_arial_layout(layout):
    """Montserrat thiếu nhiều dấu Việt — chuyển về Arial (giống app)."""
    if not isinstance(layout, dict):
        return
    f = layout.get('font')
    if isinstance(f, dict) and str(f.get('name') or '').lower() == 'montserrat':
        f['name'] = 'Arial'
        f['path'] = ''
    for k in ('title', 'info', 'channel'):
        c = layout.get(k)
        if isinstance(c, dict) and str(c.get('font') or '').lower() == 'montserrat':
            c['font'] = 'Arial'
            c['font_path'] = ''


def default_opts():
    """Tuỳ chọn mặc định — trùng giá trị mặc định của app desktop."""
    return {
        'n_per_slide': 0,
        'slides_per_file': 0,
        'ar_label': '4:3',
        'slide_size': '16:9',
        'channel_filter': ALL_CHANNELS,
        'city_filter': ALL_CITIES,
        'district_filter': ALL_DISTRICTS,
        'sort_mode': 'city',
        'pad_blank_slides': False,
        'split_export_by': 'none',
        'channel_enabled': True,
        'channel_template': 'Report {channel} 2026',
        'write_checklist': True,
        'store_index': True,
        'qty_show': {'LCD': True, 'DP': True, 'DS': True, 'GP': True},
        'visible': {n: True for n in ELEMENT_NAMES},
        'fx': dict(DEFAULT_FX),
        'slide_style': 'report',
        'scan_subfolders': True,
    }


def preset_template(name='Báo cáo'):
    """{'layout', 'opts'} của mẫu có sẵn (Báo cáo / SALESKIT / THE NELSON)."""
    pack = builtin_pack(name if name in PRESET_NAMES else 'Báo cáo')
    opts = default_opts()
    opts['slide_style'] = pack['slide_style']
    opts['visible'].update(pack['visible'])
    opts['channel_enabled'] = pack['channel_enabled']
    if name in ('SALESKIT', 'THE NELSON'):
        opts['dept'] = 'bd'
    return {'layout': copy.deepcopy(pack['layout']), 'opts': opts}


# Các tuỳ chọn được phép mang từ app / trang web sang (bỏ đường dẫn máy người dùng).
PORTABLE_OPTS = (
    'n_per_slide', 'slides_per_file', 'ar_label', 'slide_size', 'channel_filter',
    'city_filter', 'district_filter', 'sort_mode', 'pad_blank_slides', 'split_export_by',
    'channel_enabled', 'channel_template', 'write_checklist', 'store_index', 'qty_show',
    'visible', 'fx', 'slide_style', 'scan_subfolders', 'dept',
)
# Hiệu ứng cần file trên máy người dùng: bỏ khi chạy trên máy chủ.
_LOCAL_FX = ('logo_path', 'location_list')


def portable_preset(layout, opts):
    """Chụp bố cục + tuỳ chọn để lưu lên máy chủ (không kèm đường dẫn máy)."""
    o = {k: copy.deepcopy(opts[k]) for k in PORTABLE_OPTS if k in (opts or {})}
    fx = dict(o.get('fx') or {})
    for k in _LOCAL_FX:
        fx.pop(k, None)
    if fx.get('logo_enable'):
        fx['logo_enable'] = False
    o['fx'] = fx
    lay = copy.deepcopy(layout or {})
    f = lay.get('font')
    if isinstance(f, dict):
        f['path'] = ''          # đường dẫn .ttf trên máy người dùng: máy chủ tự tìm theo tên
    for k in ('title', 'info', 'channel'):
        if isinstance(lay.get(k), dict):
            lay[k]['font_path'] = ''
    return {'layout': lay, 'opts': o}


def resolve_preset(preset):
    """Chuẩn hoá preset {'layout','opts'} → (layout đủ khoá, opts đủ khoá)."""
    preset = preset or {}
    layout = deep_merge(DEFAULT_LAYOUT, preset.get('layout') or {})
    use_arial_layout(layout)
    opts = default_opts()
    for k, v in (preset.get('opts') or {}).items():
        if k in ('visible', 'qty_show', 'fx') and isinstance(v, dict):
            opts[k] = {**opts[k], **v}
        elif k in PORTABLE_OPTS:
            opts[k] = copy.deepcopy(v)
    opts['n_per_slide'] = G.parse_n_per_slide(opts.get('n_per_slide', 0))
    opts['slides_per_file'] = G.parse_slides_per_file(opts.get('slides_per_file', 0))
    if opts.get('slide_style') not in ('report', 'saleskit', 'nelson'):
        opts['slide_style'] = 'report'
    return layout, opts


def excel_fields(excel_path):
    """Kênh / tỉnh / quận có trong FILE TỔNG — để trang web dựng ô lọc."""
    src = ExcelSource()
    src.load(excel_path)
    rows = list(src.rows or [])
    return {
        'rows': len(rows),
        'channels': src.channels() if rows else [],
        'cities': geo_field_values(rows, 'City'),
        'districts': geo_field_values(rows, 'District'),
        'error': '' if rows else (getattr(src, 'error', '') or 'Không đọc được FILE TỔNG.'),
    }


class Sources:
    """FILE TỔNG + kho avatar đã đọc sẵn — máy chủ giữ trong bộ nhớ để phân tích nhanh."""

    def __init__(self, excel_path, avatar_dir=None):
        self.excel = ExcelSource()
        if not self.excel.load(excel_path) and not self.excel.rows:
            raise ValueError(getattr(self.excel, 'error', '') or 'Không đọc được FILE TỔNG.')
        self.by_code = self.excel.by_code(None, include_off=True)
        self.merged = build_merged_groups(self.by_code)
        self.avatar_dir = avatar_dir
        self.avatars = AvatarIndex()
        if avatar_dir and os.path.isdir(avatar_dir):
            self.avatars.set_folder(avatar_dir)
        self._scope = {}

    def scope(self, channel):
        if channel == ALL_CHANNELS or not self.by_code:
            return self.by_code, self.merged
        if channel not in self._scope:
            bc = scope_by_channel(self.by_code, channel, ALL_CHANNELS)
            self._scope[channel] = (bc, build_merged_groups(bc))
        return self._scope[channel]

    def avatar_for(self, code):
        row, mcode, kind = match_row(code, self.by_code, self.merged)
        excel_code = None
        if kind in ('exact', 'merged'):
            excel_code = mcode or str((row or {}).get('Code_RP') or '')
        return self.avatars.resolve(code, excel_code)


def group_names(rels):
    """Nhóm ảnh theo mã chỉ từ TÊN file (giống ImageLibrary.groups, không cần đọc đĩa)."""
    from .datasource import image_group_code
    from .constants import IMG_EXTS
    g = OrderedDict()
    seen = set()
    items = [r for r in (rels or []) if str(r).lower().endswith(IMG_EXTS)]
    for rel in sorted(items, key=lambda p: os.path.basename(str(p).replace('\\', '/')).casefold()):
        if rel in seen:
            continue
        seen.add(rel)
        code = image_group_code(rel)
        if code:
            g.setdefault(code, []).append(rel)
    return g


def analyze(src, rels, preset=None):
    """Bản xem trước cho trang biên tập web: nhóm, thứ tự, dòng Excel, bảng, avatar.

    `rels` là đường dẫn tương đối ảnh trên máy người dùng (chưa tải lên).
    Trả dict JSON được — cùng bước lọc/xếp như prepare() để thứ tự slide khớp.
    """
    from .datasource import (clean, build_info_rows, build_info_table, build_nelson_table,
                             channel_text, _address_line, screen_qty)
    _layout, opts = resolve_preset(preset)
    groups_all = group_names(rels)
    ch = opts.get('channel_filter', ALL_CHANNELS)
    by_code_scope, merged_scope = src.scope(ch)
    filtered = filter_groups_by_geo(
        groups_all, src.by_code, src.merged,
        city=opts.get('city_filter', ALL_CITIES),
        district=opts.get('district_filter', ALL_DISTRICTS),
        channel=ch, all_cities=ALL_CITIES, all_districts=ALL_DISTRICTS,
        all_channels=ALL_CHANNELS)
    ordered = order_groups(filtered, src.by_code, src.merged, src.excel.rows,
                           opts.get('sort_mode', 'city'), order_rows=list(src.excel.rows or []))
    qty = dict(opts['qty_show'])
    style = opts.get('slide_style') or 'report'
    tpl = opts.get('channel_template') or 'Report {channel} 2026'
    base = os.path.abspath(src.avatar_dir) if src.avatar_dir else ''
    out = []
    counts = {'exact': 0, 'merged': 0, 'fuzzy': 0, 'none': 0, 'avatar': 0}
    for code, files in ordered.items():
        row, mcode, kind = match_row(code, src.by_code, src.merged)
        row = row or {}
        counts[kind if kind in counts else 'none'] += 1
        sibs = [r for r in (src.excel.rows or [])
                if str(r.get('Code_RP') or '').strip().upper()
                == str(row.get('Code_RP') or mcode or code or '').upper()]
        av = src.avatar_for(code) if opts['visible'].get('avatar', True) else None
        av_rel = ''
        if av and base:
            av_rel = os.path.relpath(av, base).replace(os.sep, '/')
            counts['avatar'] += 1
        sq = screen_qty(row, qty) if opts.get('pad_blank_slides') else 0
        out.append({
            'code': code, 'files': files, 'kind': kind, 'mcode': mcode or '',
            'name': str(clean(row.get('Name'))).strip() or code,
            'channel': str(clean(row.get('Channel'))).strip(),
            'city': str(clean(row.get('City'))).strip(),
            'district': str(clean(row.get('District'))).strip(),
            'info_rows': [[l, str(v), bool(a)] for (l, v, a) in
                          build_info_rows(row, len(files) or sq or 1, qty)] if row else [],
            'info_table': build_info_table(row, sibs or [row], qty) if (row and style == 'saleskit') else None,
            'nelson_table': build_nelson_table(row, sibs or [row], qty) if (row and style == 'nelson') else None,
            'channel_text': (_address_line(row) if style == 'nelson' else channel_text(row, tpl)) if row else '',
            'avatar': av_rel,
            'slides': G.slides_for_group(files, opts['n_per_slide'], sq,
                                         bool(opts.get('pad_blank_slides'))),
        })
    hidden = [c for c in groups_all if c not in ordered]
    return {
        'groups': out,
        'photos': sum(len(v) for v in groups_all.values()),
        'groups_all': len(groups_all),
        'filtered_out': len(hidden),
        'counts': counts,
        'slides': sum(g['slides'] for g in out),
        'off_codes': len(getattr(src.excel, 'off_codes', None) or ()),
    }


class Prepared:
    """Kết quả chuẩn bị: job xuất + báo cáo đối chiếu (QA) + vài số liệu."""

    def __init__(self, job, qa, stats):
        self.job = job
        self.qa = qa
        self.stats = stats


def prepare(*, excel_path, photo_dirs, out_path, preset=None, avatar_dir=None,
            order_list_path=None, bg_image=None):
    """Đọc FILE TỔNG + ảnh → ExportJob giống hệt app desktop."""
    layout, opts = resolve_preset(preset)

    excel = ExcelSource()
    if not excel.load(excel_path) and not excel.rows:
        raise ValueError(getattr(excel, 'error', '') or 'Không đọc được FILE TỔNG.')
    order_list = ExcelSource()
    if order_list_path and os.path.isfile(order_list_path):
        order_list.load(order_list_path)

    by_code = excel.by_code(None, include_off=True)
    merged = build_merged_groups(by_code)
    ch = opts.get('channel_filter', ALL_CHANNELS)
    if ch == ALL_CHANNELS or not by_code:
        by_code_scope, merged_scope = by_code, merged
    else:
        by_code_scope = scope_by_channel(by_code, ch, ALL_CHANNELS)
        merged_scope = build_merged_groups(by_code_scope)

    imglib = ImageLibrary()
    imglib.set_folders([d for d in (photo_dirs or []) if d and os.path.isdir(d)],
                       recursive=bool(opts.get('scan_subfolders', True)))
    groups_all = OrderedDict(imglib.groups())
    if not groups_all:
        raise ValueError('Không thấy ảnh nào trong thư mục đã tải lên.')

    filtered = filter_groups_by_geo(
        groups_all, by_code, merged,
        city=opts.get('city_filter', ALL_CITIES),
        district=opts.get('district_filter', ALL_DISTRICTS),
        channel=ch, all_cities=ALL_CITIES, all_districts=ALL_DISTRICTS,
        all_channels=ALL_CHANNELS)
    sort_rows = list(order_list.rows or []) or list(excel.rows or [])
    groups = OrderedDict((code, list(files)) for code, files in order_groups(
        filtered, by_code, merged, excel.rows, opts.get('sort_mode', 'city'),
        order_rows=sort_rows).items())
    if not groups:
        raise ValueError('Không còn nhóm ảnh nào sau khi lọc kênh / tỉnh / quận.')

    avatars = AvatarIndex()
    if avatar_dir and os.path.isdir(avatar_dir):
        avatars.set_folder(avatar_dir)
    avatar_map = {}
    if opts['visible'].get('avatar', True):
        for code in groups:
            row, mcode, kind = match_row(code, by_code, merged)
            excel_code = None
            if kind in ('exact', 'merged'):
                excel_code = mcode or str((row or {}).get('Code_RP') or '')
            p = avatars.resolve(code, excel_code)
            if p:
                avatar_map[code.upper()] = p

    qa = QA.check(OrderedDict(groups_all), by_code_scope, merged_scope)
    slide_w, slide_h = slide_inches(opts.get('slide_size', '16:9'))
    job = exporter.ExportJob(
        layout=layout,
        groups=groups,
        excel_by_code=dict(by_code),
        excel_rows=list(excel.rows) + list(getattr(excel, 'off_rows', None) or []),
        avatar_map=avatar_map,
        out_path=out_path,
        n_per_slide=opts['n_per_slide'],
        slides_per_file=opts['slides_per_file'],
        img_ar=AR_CHOICES.get(opts.get('ar_label', '4:3'), 4 / 3),
        bg_image=bg_image if bg_image and os.path.isfile(bg_image) else None,
        channel_enabled=bool(opts.get('channel_enabled', True)),
        channel_template=opts.get('channel_template') or 'Report {channel} 2026',
        visible=dict(opts['visible']),
        fx=copy.deepcopy(opts['fx']),
        slide_style=opts['slide_style'],
        pad_blank_slides=bool(opts.get('pad_blank_slides', False)),
        sort_mode=opts.get('sort_mode', 'city'),
        sort_rows=sort_rows,
        city_filter=opts.get('city_filter', ALL_CITIES),
        district_filter=opts.get('district_filter', ALL_DISTRICTS),
        split_export_by=opts.get('split_export_by', 'none'),
        export_pdf=False,
        qty_show=dict(opts['qty_show']),
        store_index=bool(opts.get('store_index', True)),
        excel_scope_by_code=dict(by_code_scope or by_code),
        slide_w=slide_w,
        slide_h=slide_h,
    )
    n_ov = 0
    if opts['fx'].get('dashboard', True):
        try:
            from .datasource import build_overview
            n_ov = len(build_overview(groups, by_code_scope, merged_scope,
                                      job.qty_show)['rows'])
        except Exception:
            n_ov = 0
    ng, ns, nf = exporter.estimate(
        groups, job.n_per_slide, overview_rows=n_ov, slides_per_file=job.slides_per_file,
        by_code=by_code, merged=merged, pad_blank=job.pad_blank_slides,
        store_index=job.store_index, qty_show=job.qty_show)
    stats = {
        'photos': sum(len(v) for v in groups_all.values()),
        'groups_all': len(groups_all),
        'groups': ng,
        'slides_est': ns,
        'files_est': nf,
        'avatars': len(avatar_map),
        'write_checklist': bool(opts.get('write_checklist', True)),
    }
    return Prepared(job, qa, stats)


def run(prepared, progress_cb=None, log_cb=None, cancel=None):
    """Xuất PPTX (+ checklist Excel như app). Trả (ExportReport, đường dẫn checklist)."""
    rep = exporter.export_pptx(prepared.job, progress_cb, log_cb, cancel)
    checklist = ''
    if (prepared.stats.get('write_checklist') and not rep.cancelled
            and getattr(prepared.qa, 'has_excel', False)):
        try:
            checklist = QA.write_checklist(QA.checklist_path(prepared.job.out_path), prepared.qa)
        except Exception as e:
            if log_cb:
                log_cb(f'Không ghi được checklist: {e}')
    return rep, checklist
