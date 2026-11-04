#!/usr/bin/env python3
"""Chạy MỘT lần xuất PPTX trên máy chủ — tiến trình riêng do autopptx_server.py gọi.

    python3 render_job.py JOB_DIR

JOB_DIR chứa job.json (do máy chủ ghi), photos/, có thể có order.xlsx.
Ghi tiến độ vào status.json (atomic); dừng khi thấy file CANCEL.
Kết quả: JOB_DIR/out/*.pptx, *_list_cuahang.xlsx, *_checklist.xlsx
"""
import json
import os
import sys
import threading
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
for d in (HERE, os.path.dirname(HERE)):
    if os.path.isdir(os.path.join(d, 'autopptx2')) and d not in sys.path:
        sys.path.insert(0, d)


def main(job_dir):
    with open(os.path.join(job_dir, 'job.json'), encoding='utf-8') as f:
        spec = json.load(f)
    status_path = os.path.join(job_dir, 'status.json')
    st = {'state': 'running', 'done': 0, 'total': 0, 'msg': 'Đang đọc FILE TỔNG và ảnh…',
          'log': [], 'started': time.time()}
    lock = threading.Lock()
    last_write = [0.0]

    def write(force=False):
        t = time.time()
        if not force and t - last_write[0] < 0.7:
            return
        last_write[0] = t
        with lock:
            tmp = status_path + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(st, f, ensure_ascii=False)
            # Windows: máy chủ đang đọc status.json thì không ghi đè được — thử lại.
            for i in range(40 if force else 3):
                try:
                    os.replace(tmp, status_path)
                    break
                except PermissionError:
                    time.sleep(0.05)

    def log(msg):
        with lock:
            st['log'] = (st['log'] + [str(msg)])[-40:]
        write()

    cancel = threading.Event()

    def watch_cancel():
        while not cancel.is_set():
            if os.path.exists(os.path.join(job_dir, 'CANCEL')):
                cancel.set()
                return
            time.sleep(0.5)

    threading.Thread(target=watch_cancel, daemon=True).start()
    write(True)
    try:
        from autopptx2 import pipeline as P
        out_dir = os.path.join(job_dir, 'out')
        os.makedirs(out_dir, exist_ok=True)
        name = spec.get('out_name') or 'Report.pptx'
        order = os.path.join(job_dir, 'order.xlsx')
        prep = P.prepare(
            excel_path=os.path.join(job_dir, 'excel.xlsx'),
            photo_dirs=[os.path.join(job_dir, 'photos')],
            out_path=os.path.join(out_dir, name),
            preset=spec.get('preset'),
            avatar_dir=spec.get('avatar_dir'),
            order_list_path=order if os.path.isfile(order) else None,
            bg_image=spec.get('bg_image'))
        st['stats'] = prep.stats
        st['total'] = prep.stats.get('slides_est', 0)
        st['msg'] = f"{prep.stats['groups']} nhóm ảnh → khoảng {prep.stats['slides_est']} slide"
        write(True)

        def prog(done, total):
            st['done'], st['total'] = done, total
            st['msg'] = f'Đang dựng slide {done}/{total}'
            write()

        rep, checklist = P.run(prep, prog, log, cancel)
        files = [f for f in rep.files if f and os.path.isfile(f)]
        if checklist and os.path.isfile(checklist):
            files.append(checklist)
        st['files'] = [{'name': os.path.basename(f), 'size': os.path.getsize(f)} for f in files]
        st['report'] = {
            'slides': rep.slides, 'groups': rep.groups, 'seconds': round(rep.seconds, 1),
            'unmatched': list(rep.unmatched)[:500],
            'fuzzy': [list(x) for x in rep.fuzzy][:500],
            'missing_avatars': list(rep.missing_avatars)[:500],
        }
        if rep.cancelled:
            st['state'], st['msg'] = 'cancelled', 'Đã huỷ.'
        else:
            st['state'] = 'done'
            st['msg'] = f'Xong: {rep.slides} slide trong {rep.seconds:.0f} giây.'
    except Exception as e:
        traceback.print_exc()
        st['state'] = 'failed'
        st['msg'] = str(e) if isinstance(e, ValueError) else f'Lỗi khi xuất: {type(e).__name__}: {e}'
    st['finished'] = time.time()
    write(True)
    return 0 if st['state'] == 'done' else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
