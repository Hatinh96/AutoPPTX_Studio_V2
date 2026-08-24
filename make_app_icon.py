"""Tạo app_icon.ico (Windows) và app_icon.icns (macOS dock) từ logo G vuông, đủ size Retina."""
import os
import sys

from PIL import Image

ROOT = os.path.dirname(os.path.abspath(__file__))
SOURCES = [
    os.path.join(ROOT, 'LOGO', 'Logo Golden Asia_Symbol_Standard.png'),
    os.path.join(ROOT, 'LOGO', 'Logo Golden_Symbol Premium.png'),
    os.path.join(ROOT, 'LOGO', 'Logo Golden_Symbol Multicolor.png'),
]
ICO_SIZES = (16, 24, 32, 48, 64, 128, 256)
ICNS_SIZES = (16, 32, 64, 128, 256, 512, 1024)
PAD = 0.16
BG = (255, 255, 255, 255)


def _src():
    for p in SOURCES:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError('Không thấy file logo trong LOGO/')


def square_icon(src, size, pad=PAD, bg=BG):
    im = Image.open(src).convert('RGBA')
    canvas = Image.new('RGBA', (size, size), bg)
    inner = max(1, int(size * (1.0 - 2.0 * pad)))
    ratio = min(inner / max(1, im.width), inner / max(1, im.height))
    nw = max(1, int(im.width * ratio))
    nh = max(1, int(im.height * ratio))
    logo = im.resize((nw, nh), Image.Resampling.LANCZOS)
    x = (size - nw) // 2
    y = (size - nh) // 2
    canvas.paste(logo, (x, y), logo)
    return canvas


def write_icons(root=ROOT):
    src = _src()
    master = square_icon(src, 1024)
    ico_path = os.path.join(root, 'app_icon.ico')
    icns_path = os.path.join(root, 'app_icon.icns')
    ico_master = square_icon(src, 256)
    ico_master.save(ico_path, format='ICO', sizes=[(s, s) for s in ICO_SIZES])
    try:
        master.save(icns_path, format='ICNS')
    except Exception as e:
        print('Không ghi được .icns (Pillow/platform):', e)
        icns_path = None
    png_path = os.path.join(root, 'app_icon_1024.png')
    master.save(png_path, format='PNG')
    return ico_path, icns_path


if __name__ == '__main__':
    ico, icns = write_icons()
    print('Wrote', ico)
    if icns:
        print('Wrote', icns)
    sys.exit(0)
