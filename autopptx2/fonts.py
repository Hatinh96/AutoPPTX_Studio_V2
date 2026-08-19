"""Danh mục font Windows — không phụ thuộc Tkinter.

Preview (PIL) và PPTX dùng cùng family + cùng file .ttf khi có trên máy.
"""
import os
import json
import threading
from functools import lru_cache

from PIL import Image, ImageDraw, ImageFont

from .constants import CONFIG_DIR

CACHE_PATH = os.path.join(CONFIG_DIR, "fonts_cache.json")
VIET_SAMPLE = "ÁÀẢÃẠăâêôơưĐ"
PINNED = ("Be Vietnam Pro", "Montserrat", "Arial", "Segoe UI", "Calibri",
          "Tahoma", "Times New Roman")

_lock = threading.Lock()
_catalog = None          # list[dict]
_by_family = None        # {family.lower(): {regular, bold, italic, bolditalic}}


def font_dirs():
    dirs = []
    windir = os.environ.get("WINDIR", r"C:\Windows")
    dirs.append(os.path.join(windir, "Fonts"))
    local = os.environ.get("LOCALAPPDATA") or ""
    if local:
        dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
    return [d for d in dirs if os.path.isdir(d)]


def _dirs_sig():
    out = []
    for d in font_dirs():
        try:
            out.append((d, os.path.getmtime(d)))
        except Exception:
            out.append((d, None))
    return out


def _style_key(style):
    s = (style or "").lower()
    bold = "bold" in s or "bd" in s or s.endswith("b") or "black" in s or "heavy" in s
    italic = "italic" in s or "oblique" in s or "it" == s or s.endswith("i")
    if bold and italic:
        return "bolditalic"
    if bold:
        return "bold"
    if italic:
        return "italic"
    return "regular"


def _has_viet(font):
    try:
        for ch in VIET_SAMPLE:
            bb = font.getbbox(ch)
            if (bb[2] - bb[0]) < 2:
                return False
        return True
    except Exception:
        return False


def _scan():
    seen_files = set()
    items = []
    families = {}
    for folder in font_dirs():
        try:
            names = os.listdir(folder)
        except Exception:
            continue
        for fn in names:
            ext = os.path.splitext(fn)[1].lower()
            if ext not in (".ttf", ".otf", ".ttc"):
                continue
            path = os.path.join(folder, fn)
            if path in seen_files:
                continue
            seen_files.add(path)
            max_idx = 8 if ext == ".ttc" else 1
            for idx in range(max_idx):
                try:
                    fnt = ImageFont.truetype(path, 22, index=idx)
                except Exception:
                    break
                try:
                    family, style = fnt.getname()
                except Exception:
                    family, style = os.path.splitext(fn)[0], "Regular"
                family = (family or "").strip()
                if not family:
                    continue
                rec = {
                    "family": family,
                    "style": style or "Regular",
                    "path": path,
                    "index": idx,
                    "viet": _has_viet(fnt),
                    "key": _style_key(style),
                }
                items.append(rec)
                slot = families.setdefault(family.lower(), {
                    "family": family, "regular": None, "bold": None,
                    "italic": None, "bolditalic": None, "viet": False,
                })
                slot["viet"] = slot["viet"] or rec["viet"]
                if slot[rec["key"]] is None:
                    slot[rec["key"]] = rec
                if rec["key"] == "regular":
                    slot["family"] = family
    # family list: pinned first, then A–Z, one row per family (prefer Regular)
    rows = []
    used = set()
    for pin in PINNED:
        slot = families.get(pin.lower())
        if slot:
            rows.append(_family_row(slot))
            used.add(pin.lower())
    rest = sorted((s for k, s in families.items() if k not in used),
                  key=lambda s: s["family"].casefold())
    rows.extend(_family_row(s) for s in rest)
    return {"sig": _dirs_sig(), "families": rows, "all": items}


def _family_row(slot):
    rec = (slot["regular"] or slot["bold"] or slot["italic"]
           or slot["bolditalic"] or {})
    return {
        "family": slot["family"],
        "path": rec.get("path", ""),
        "index": rec.get("index", 0),
        "viet": bool(slot["viet"]),
        "regular": (slot["regular"] or {}).get("path"),
        "bold": (slot["bold"] or {}).get("path"),
        "italic": (slot["italic"] or {}).get("path"),
        "bolditalic": (slot["bolditalic"] or {}).get("path"),
        "regular_index": (slot["regular"] or rec).get("index", 0),
        "bold_index": (slot["bold"] or rec).get("index", 0),
        "italic_index": (slot["italic"] or rec).get("index", 0),
        "bolditalic_index": (slot["bolditalic"] or rec).get("index", 0),
    }


def _sig_key():
    return [[d, mt] for d, mt in _dirs_sig()]


def _load_cache():
    try:
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("sig") == _sig_key() and data.get("families"):
            return data
    except Exception:
        return None
    return None


def _save_cache(data):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        dump = dict(data)
        dump["sig"] = _sig_key()
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(dump, f)
    except Exception:
        pass


def catalog(force=False):
    """Danh sách family (mỗi family 1 dòng). Quét đĩa lần đầu, sau đó cache."""
    global _catalog, _by_family
    with _lock:
        if _catalog is not None and not force:
            return _catalog
        data = None if force else _load_cache()
        if not data or not data.get("families"):
            data = _scan()
            _save_cache(data)
        _catalog = data["families"]
        _by_family = {r["family"].lower(): r for r in _catalog}
        return _catalog


def ready():
    return _catalog is not None


def family_info(name):
    if _by_family is None:
        return None
    return _by_family.get((name or "").lower())


def preview_sample(path, index=0, text="ÁÀẢÃẠ", size=(392, 54)):
    """Ảnh mẫu từ đúng file .ttf — cùng nguồn với preview slide / PPTX."""
    bg, fg = (28, 28, 32), (245, 245, 247)
    im = Image.new("RGB", size, bg)
    draw = ImageDraw.Draw(im)
    fnt = None
    if path and os.path.isfile(path):
        try:
            fnt = ImageFont.truetype(path, 28, index=int(index or 0))
        except Exception:
            fnt = None
    if fnt is None:
        try:
            fnt = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 28)
        except Exception:
            fnt = ImageFont.load_default()
    draw.text((12, 10), text, font=fnt, fill=fg)
    return im


def resolve_font_file(name, bold=False, italic=False):
    """(path, index) của file .ttf/.otf khớp family + đậm/nghiêng."""
    info = family_info(name)
    if not info:
        return None, 0
    if bold and italic:
        order = ("bolditalic", "bold", "italic", "regular")
    elif bold:
        order = ("bold", "regular", "bolditalic", "italic")
    elif italic:
        order = ("italic", "regular", "bolditalic", "bold")
    else:
        order = ("regular", "bold", "italic", "bolditalic")
    for k in order:
        p = info.get(k)
        if p and os.path.isfile(p):
            return p, int(info.get(k + "_index", 0) or 0)
    p = info.get("path")
    if p and os.path.isfile(p):
        return p, int(info.get("index", 0) or 0)
    return None, 0


def search_families(query, limit=80):
    q = (query or "").strip().casefold()
    rows = catalog()
    if not q:
        return rows[:limit]
    hits = [r for r in rows if q in r["family"].casefold()]
    return hits[:limit]


@lru_cache(maxsize=96)
def pil_font(path, size, index=0):
    size = max(7, int(size))
    try:
        return ImageFont.truetype(path, size, index=index)
    except Exception:
        return ImageFont.load_default()
