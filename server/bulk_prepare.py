"""Chuẩn bị đưa hàng loạt avatar / ảnh nền lên máy chủ mà không phải tải lên bản gốc.

Thu nhỏ ảnh ngay trên máy (cùng quy tắc với máy chủ), ghi vào thư mục tạm kèm
src_map.json (sha ảnh gốc). Sau đó chép thư mục tạm lên VPS và chạy
`autopptx_server.py reindex src_map.json` (xem server/README.md).

    python server/bulk_prepare.py avatars "Data avatar" OUT_DIR
    python server/bulk_prepare.py backgrounds "Data background/Background" OUT_DIR --prefix shared
"""
import argparse
import importlib.util
import json
import os
import sys
import unicodedata
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))


def load_server():
    os.environ.setdefault('AP_DATA', os.path.join(HERE, '_unused_data'))
    spec = importlib.util.spec_from_file_location('ap_srv', os.path.join(HERE, 'autopptx_server.py'))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('kind', choices=('avatars', 'backgrounds'))
    ap.add_argument('src')
    ap.add_argument('out')
    ap.add_argument('--prefix', default='', help='thư mục con phía máy chủ (nền chung: shared)')
    a = ap.parse_args()
    S = load_server()
    out_kind = os.path.join(a.out, a.kind)
    jobs, skipped = [], []
    for root, _ds, fns in os.walk(a.src):
        for fn in fns:
            full = os.path.join(root, fn)
            rel = os.path.relpath(full, a.src).replace(os.sep, '/')
            if a.prefix:
                rel = a.prefix.strip('/') + '/' + rel
            rel = S.clean_rel(unicodedata.normalize('NFC', rel), a.kind)
            if not rel:
                if fn.lower().endswith(S.IMG_EXTS):
                    skipped.append(os.path.relpath(full, a.src))
                continue
            jobs.append((full, rel))
    src_map, stats = {}, {'in': 0, 'out': 0}

    def one(job):
        full, rel = job
        with open(full, 'rb') as f:
            data = f.read()
        sha = S.sha256_bytes(data)
        out = S.shrink_avatar(data, os.path.splitext(rel)[1]) if a.kind == 'avatars' else data
        dst = os.path.join(out_kind, *rel.split('/'))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with open(dst, 'wb') as f:
            f.write(out)
        return rel, sha, len(data), len(out)

    with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as ex:
        for i, (rel, sha, n_in, n_out) in enumerate(ex.map(one, jobs), 1):
            src_map[rel] = sha
            stats['in'] += n_in
            stats['out'] += n_out
            if i % 500 == 0:
                print(f'{i}/{len(jobs)}', flush=True)
    map_path = os.path.join(a.out, f'src_map_{a.kind}.json')
    with open(map_path, 'w', encoding='utf-8') as f:
        json.dump({a.kind: src_map}, f, ensure_ascii=False)
    print(json.dumps({'files': len(src_map), 'mb_in': round(stats['in'] / 1048576, 1),
                      'mb_out': round(stats['out'] / 1048576, 1), 'skipped': skipped[:20],
                      'skipped_count': len(skipped), 'map': map_path}, ensure_ascii=False))


if __name__ == '__main__':
    sys.exit(main())
