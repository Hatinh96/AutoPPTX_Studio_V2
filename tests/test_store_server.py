"""Kiểm tra máy chủ kho tài nguyên (server/autopptx_server.py) chạy thật trên cổng tạm."""
import importlib.util
import io
import json
import os
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVER_FILE = os.path.join(ROOT, 'server', 'autopptx_server.py')


def load_server(data_dir):
    os.environ['AP_DATA'] = data_dir
    os.environ['AP_PORT'] = '0'
    os.environ['AP_AVATAR_MAX_PX'] = '400'
    spec = importlib.util.spec_from_file_location('autopptx_server_test', SERVER_FILE)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def make_jpeg(w, h, color=(200, 30, 30)):
    from PIL import Image
    b = io.BytesIO()
    Image.new('RGB', (w, h), color).save(b, 'JPEG', quality=95)
    return b.getvalue()


def make_xlsx(codes):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'LIST'
    ws.append(['STT', 'Code_RP', 'Name', 'Address', 'District', 'City', 'Channel', 'Qty'])
    for i, c in enumerate(codes, 1):
        ws.append([i, c, f'Điểm {c}', f'{i} Lê Lợi', 'Quận 1', 'Hồ Chí Minh', 'Building', 1])
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


class StoreServerTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.S = load_server(cls.tmp.name)
        cls.S.init_db()
        con = cls.S.db()
        now = cls.S.now_iso()
        con.execute("INSERT INTO users (email, name, role, pass_hash, must_change, created_at) "
                    "VALUES ('boss@x.vn','Boss','super_admin',?,0,?)",
                    (cls.S.hash_password('adminpass1'), now))
        con.execute("INSERT INTO users (email, name, role, pass_hash, must_change, created_at) "
                    "VALUES ('staff@x.vn','Staff','user',?,0,?)",
                    (cls.S.hash_password('staffpass1'), now))
        con.commit()
        con.close()
        cls.srv = cls.S.ThreadingHTTPServer(('127.0.0.1', 0), cls.S.Handler)
        cls.base = f'http://127.0.0.1:{cls.srv.server_address[1]}'
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.admin = cls.login('boss@x.vn', 'adminpass1')['token']
        cls.staff = cls.login('staff@x.vn', 'staffpass1')['token']

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        cls.tmp.cleanup()

    @classmethod
    def req(cls, method, path, token=None, body=None, headers=None, raw=False):
        h = dict(headers or {})
        if token:
            h['Authorization'] = 'Bearer ' + token
        data = body
        if isinstance(body, dict):
            data = json.dumps(body).encode()
            h['Content-Type'] = 'application/json'
        r = urllib.request.Request(cls.base + path, data=data, method=method, headers=h)
        try:
            with urllib.request.urlopen(r, timeout=30) as resp:
                b = resp.read()
                return (resp.status, b) if raw else (resp.status, json.loads(b or b'{}'))
        except urllib.error.HTTPError as e:
            b = e.read()
            try:
                return e.code, json.loads(b)
            except Exception:
                return e.code, b

    @classmethod
    def login(cls, email, pw):
        st, d = cls.req('POST', '/api/login', body={'email': email, 'password': pw})
        assert st == 200, d
        return d

    @staticmethod
    def q(path):
        return '/'.join(urllib.parse.quote(p) for p in path.split('/'))

    def test_login_wrong_password(self):
        st, d = self.req('POST', '/api/login', body={'email': 'staff@x.vn', 'password': 'nope'})
        self.assertEqual(st, 401)
        self.assertIn('Sai', d['msg'])

    def test_requires_token(self):
        st, _ = self.req('GET', '/api/manifest/avatars')
        self.assertEqual(st, 401)

    def test_avatar_upload_shrink_manifest_and_permissions(self):
        big = make_jpeg(1200, 800)
        path = 'Coffee Shop & Milktea/Quán Cà Phê/1994HAIRA.jpg'
        st, d = self.req('PUT', '/api/files/avatars/' + self.q(path), self.staff, big)
        self.assertEqual(st, 403)
        st, d = self.req('PUT', '/api/files/avatars/' + self.q(path), self.admin, big)
        self.assertEqual(st, 200, d)
        self.assertFalse(d['unchanged'])
        self.assertLess(d['size'], len(big))
        st, d = self.req('PUT', '/api/files/avatars/' + self.q(path), self.admin, big)
        self.assertTrue(d['unchanged'])
        st, m = self.req('GET', '/api/manifest/avatars', self.staff)
        row = next(f for f in m['files'] if f['path'] == path)
        self.assertEqual(row['code'], '1994HAIR')
        st, b = self.req('GET', '/api/files/avatars/' + self.q(path), self.staff, raw=True)
        self.assertEqual(st, 200)
        from PIL import Image
        self.assertEqual(max(Image.open(io.BytesIO(b)).size), 400)
        self.assertEqual(self.S.sha256_bytes(b), row['sha'])

    def test_path_traversal_rejected(self):
        for bad in ('../x.jpg', 'a/../../x.jpg', 'a/.hidden.jpg', 'a/b.exe'):
            st, _ = self.req('PUT', '/api/files/avatars/' + urllib.parse.quote(bad, safe='/'),
                             self.admin, b'x')
            self.assertEqual(st, 400, bad)

    def test_backgrounds_shared_and_private(self):
        img = make_jpeg(320, 180, (10, 10, 200))
        st, _ = self.req('PUT', '/api/files/backgrounds/shared/nen.jpg', self.staff, img)
        self.assertEqual(st, 403)
        st, _ = self.req('PUT', '/api/files/backgrounds/shared/nen.jpg', self.admin, img)
        self.assertEqual(st, 200)
        st, me = self.req('GET', '/api/me', self.staff)
        uid = me['user']['id']
        st, _ = self.req('PUT', f'/api/files/backgrounds/users/{uid}/rieng.jpg', self.staff, img)
        self.assertEqual(st, 200)
        st, _ = self.req('PUT', f'/api/files/backgrounds/users/{uid + 99}/x.jpg', self.staff, img)
        self.assertEqual(st, 403)
        st, m = self.req('GET', '/api/manifest/backgrounds', self.staff)
        paths = {f['path'] for f in m['files']}
        self.assertIn('shared/nen.jpg', paths)
        self.assertIn(f'users/{uid}/rieng.jpg', paths)
        st, m = self.req('GET', '/api/manifest/backgrounds', self.admin)
        self.assertNotIn(f'users/{uid}/rieng.jpg', {f['path'] for f in m['files']})

    def test_excel_shared_personal_and_report(self):
        for code in ('AAA01', 'BBB02'):
            self.req('PUT', f'/api/files/avatars/Building/{code}A.jpg', self.admin,
                     make_jpeg(60, 60))
        xl = make_xlsx(['AAA01', 'BBB02', 'CCC03'])
        hdr = {'X-File-Name': urllib.parse.quote('FILE TỔNG W33.xlsx')}
        st, d = self.req('PUT', '/api/excel?target=me', self.staff, xl, hdr)
        self.assertEqual(st, 200, d)
        st, d = self.req('PUT', '/api/excel?target=all', self.staff, xl, hdr)
        self.assertEqual(st, 403)
        st, d = self.req('GET', '/api/excel/meta', self.staff)
        self.assertNotEqual(d['excel']['owner'], 'shared')
        st, d = self.req('PUT', '/api/excel?target=all', self.admin, xl, hdr)
        self.assertEqual(st, 200, d)
        self.assertEqual(d['cleared_personal'], 1)
        st, d = self.req('GET', '/api/excel/meta', self.staff)
        self.assertEqual(d['excel']['owner'], 'shared')
        self.assertEqual(d['excel']['file_name'], 'FILE TỔNG W33.xlsx')
        st, b = self.req('GET', '/api/excel/file', self.staff, raw=True)
        self.assertEqual(b, xl)
        st, r = self.req('GET', '/api/admin/report', self.admin)
        self.assertEqual(st, 200, r)
        self.assertEqual(r['codes'], 3)
        self.assertEqual([m['code'] for m in r['missing']], ['CCC03'])
        st, _ = self.req('GET', '/api/admin/report', self.staff)
        self.assertEqual(st, 403)

    def test_admin_create_user_and_reset(self):
        st, d = self.req('POST', '/api/admin/users', self.admin,
                         {'email': 'New@X.vn', 'name': 'Mới', 'role': 'user'})
        self.assertEqual(st, 200, d)
        temp = d['temp_password']
        self.assertEqual(d['user']['email'], 'new@x.vn')
        tok = self.login('new@x.vn', temp)['token']
        st, d = self.req('POST', '/api/me/password', tok, {'old': temp, 'new': 'matkhaumoi9'})
        self.assertEqual(st, 200, d)
        self.login('new@x.vn', 'matkhaumoi9')
        st, d = self.req('POST', '/api/admin/users', self.staff, {'email': 'z@x.vn'})
        self.assertEqual(st, 403)

    def test_cookie_requires_csrf_header(self):
        st, d = self.req('POST', '/api/login', body={'email': 'boss@x.vn', 'password': 'adminpass1',
                                                     'cookie': True})
        self.assertIsNone(d['token'])
        # Lấy cookie bằng tay
        r = urllib.request.Request(self.base + '/api/login', method='POST',
                                   data=json.dumps({'email': 'boss@x.vn', 'password': 'adminpass1',
                                                    'cookie': True}).encode(),
                                   headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(r) as resp:
            cookie = resp.headers['Set-Cookie'].split(';')[0]
        st, _ = self.req('GET', '/api/admin/stats', headers={'Cookie': cookie})
        self.assertEqual(st, 200)
        st, _ = self.req('POST', '/api/admin/users', body={'email': 'c@x.vn'},
                         headers={'Cookie': cookie})
        self.assertEqual(st, 403)
        st, _ = self.req('POST', '/api/admin/users', body={'email': 'c@x.vn'},
                         headers={'Cookie': cookie, 'X-Requested-With': 'autopptx'})
        self.assertEqual(st, 200)

    def test_reindex_picks_up_copied_files(self):
        dst = os.path.join(self.S.STORE, 'avatars', 'Hotel', 'ZZZ9A.png')
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        from PIL import Image
        Image.new('RGB', (20, 20)).save(dst)
        res = self.S.reindex()
        self.assertGreaterEqual(res['added'], 1)
        st, m = self.req('GET', '/api/manifest/avatars', self.staff)
        self.assertIn('Hotel/ZZZ9A.png', {f['path'] for f in m['files']})

    def test_app_client_end_to_end(self):
        """CloudClient của app ↔ máy chủ thật: đăng nhập, đẩy/tải Excel, avatar, nền."""
        from unittest.mock import patch
        import autopptx2.cloud as C
        with tempfile.TemporaryDirectory() as td:
            av_dir, bg_dir, xl_dir = (os.path.join(td, n) for n in ('av', 'bg', 'xl'))
            for d in (av_dir, bg_dir, xl_dir):
                os.makedirs(d)
            src = os.path.join(td, 'src')
            os.makedirs(os.path.join(src, 'Salon', '30SHINE'))
            with open(os.path.join(src, 'Salon', '30SHINE', 'E2E01A.jpg'), 'wb') as f:
                f.write(make_jpeg(900, 600))
            with open(os.path.join(src, 'Salon', 'E2E02_Ava.png'), 'wb') as f:
                from PIL import Image
                b = io.BytesIO()
                Image.new('RGB', (50, 50), (0, 200, 0)).save(b, 'PNG')
                f.write(b.getvalue())
            xl_path = os.path.join(td, 'Master E2E.xlsx')
            with open(xl_path, 'wb') as f:
                f.write(make_xlsx(['E2E01', 'E2E02']))
            with patch.object(C, 'avatar_cache_dir', return_value=av_dir), \
                    patch.object(C, 'bg_cache_dir', return_value=bg_dir), \
                    patch.object(C, 'excel_cache_dir', return_value=xl_dir):
                admin = C.CloudClient(self.base)
                ok, msg = admin.login('boss@x.vn', 'adminpass1')
                self.assertTrue(ok, msg)
                self.assertTrue(admin.is_admin())
                r = admin.upload_avatars(src)
                self.assertTrue(r['ok'], r)
                self.assertIn('Đã đẩy 2', r['msg'])
                self.assertIn('0 không đổi', r['msg'])
                r = admin.upload_avatars(src)          # lần 2: không đổi gì
                self.assertIn('Đã đẩy 0', r['msg'])
                r = admin.upload_excel_broadcast(xl_path)
                self.assertTrue(r['ok'], r)

                staff = C.CloudClient(self.base)
                ok, msg = staff.login('staff@x.vn', 'staffpass1')
                self.assertTrue(ok, msg)
                self.assertFalse(staff.is_admin())
                self.assertFalse(staff.upload_avatars(src)['ok'])
                r = staff.sync_excel()
                self.assertTrue(r['ok'], r)
                self.assertTrue(r['fresh'])
                self.assertEqual(r['file_name'], 'Master E2E.xlsx')
                self.assertFalse(staff.sync_excel()['fresh'])
                r = staff.sync_avatars()
                self.assertTrue(r['ok'], r)
                self.assertGreaterEqual(r['downloaded'], 2)
                self.assertTrue(os.path.isfile(os.path.join(av_dir, 'Salon', '30SHINE', 'E2E01A.jpg')))
                r = staff.sync_avatars()
                self.assertEqual(r['downloaded'], 0)

                bg = os.path.join(td, 'nen rieng.jpg')
                with open(bg, 'wb') as f:
                    f.write(make_jpeg(160, 90))
                r = staff.upload_background_files([bg])
                self.assertTrue(r['ok'], r)
                self.assertTrue(r['storage_path'].startswith(f'users/{staff.user_id}/'))
                rows = staff.list_backgrounds()
                mine = [x for x in rows if x['storage_path'] == r['storage_path']]
                self.assertEqual(len(mine), 1)
                t = staff.ensure_bg_thumb(mine[0]['storage_path'], mine[0]['updated_at'])
                self.assertTrue(t['ok'], t)

                # Phiên đã nhớ dùng lại được; đăng xuất thì phiên chết.
                again = C.CloudClient(self.base)
                ok, _ = again.resume('staff@x.vn', staff.token)
                self.assertTrue(ok)
                tok = staff.token
                staff.logout()
                ok, msg = C.CloudClient(self.base).resume('staff@x.vn', tok)
                self.assertFalse(ok)
                self.assertIn('hết', msg)

                profiles = admin.fetch_profiles()
                sid = next(p['id'] for p in profiles if p['email'] == 'staff@x.vn')
                self.assertTrue(admin.set_admin(sid, 'Staff', True)['ok'])
                self.assertTrue(admin.set_admin(sid, 'Staff', False)['ok'])

    def test_web_export_job_end_to_end(self):
        """Bản web: tạo job → tải ảnh → máy chủ dựng PPTX → tải về, đúng số slide."""
        import time as _t
        from pptx import Presentation
        if not getattr(type(self), '_worker', None):
            type(self)._worker = threading.Thread(target=self.S.job_worker, daemon=True)
            type(self)._worker.start()
        hdr = {'X-File-Name': urllib.parse.quote('Web.xlsx')}
        st, d = self.req('PUT', '/api/excel?target=all', self.admin,
                         make_xlsx(['WEB01', 'WEB02', 'WEB03']), hdr)
        self.assertEqual(st, 200, d)
        st, d = self.req('GET', '/api/excel/fields', self.staff)
        self.assertEqual(st, 200, d)
        self.assertIn('Building', d['channels'])
        st, d = self.req('GET', '/api/presets', self.staff)
        self.assertIn('builtin:SALESKIT', [p['id'] for p in d['presets']])
        # Mẫu riêng do app gửi lên (CloudClient.save_preset)
        import autopptx2.cloud as C
        cli = C.CloudClient(self.base)
        self.assertTrue(cli.login('staff@x.vn', 'staffpass1')[0])
        from autopptx2 import pipeline as P
        tpl = P.preset_template('Báo cáo')
        r = cli.save_preset('Mẫu thử', tpl['layout'], {**tpl['opts'], 'n_per_slide': 2})
        self.assertTrue(r['ok'], r)
        st, d = self.req('GET', '/api/presets', self.staff)
        mine = next(p for p in d['presets'] if p['name'] == 'Mẫu thử')
        self.assertEqual(mine['kind'], 'mine')

        st, job = self.req('POST', '/api/jobs', self.staff, {
            'preset_id': mine['id'], 'title': 'Báo cáo thử W1',
            'overrides': {'fx': {'use_timestamp': True}, 'store_index': True}})
        self.assertEqual(st, 200, job)
        jid = job['id']
        photos = {'f1/Anh/WEB01 (1).jpg': (200, 0, 0), 'f1/Anh/WEB01 (2).jpg': (0, 200, 0),
                  'f1/Anh/WEB01 (3).jpg': (0, 0, 200), 'f1/Anh/WEB02 (1).jpg': (90, 90, 0),
                  'f1/Anh/LAMA99 (1).jpg': (9, 9, 9)}
        for rel, col in photos.items():
            st, d = self.req('PUT', f'/api/jobs/{jid}/photos/' + self.q(rel), self.staff,
                             make_jpeg(640, 480, col))
            self.assertEqual(st, 200, d)
        # token sai không xem được job
        st, _ = self.req('GET', f'/api/jobs/{jid}', 'token-khong-hop-le')
        self.assertEqual(st, 401)
        st, d = self.req('POST', f'/api/jobs/{jid}/start', self.staff, {})
        self.assertEqual(st, 200, d)
        st, d = self.req('PUT', f'/api/jobs/{jid}/photos/f1/x.jpg', self.staff, make_jpeg(10, 10))
        self.assertEqual(st, 400)
        t0 = _t.time()
        while _t.time() - t0 < 180:
            st, d = self.req('GET', f'/api/jobs/{jid}', self.staff)
            if d['job']['state'] in ('done', 'failed', 'cancelled'):
                break
            _t.sleep(0.5)
        j = d['job']
        self.assertEqual(j['state'], 'done', j)
        names = [f['name'] for f in j['files']]
        self.assertIn('Báo cáo thử W1.pptx', names)
        self.assertIn('LAMA99', j['report']['unmatched'])
        st, b = self.req('GET', f'/api/jobs/{jid}/files/' + urllib.parse.quote('Báo cáo thử W1.pptx'),
                         self.staff, raw=True)
        self.assertEqual(st, 200)
        prs = Presentation(io.BytesIO(b))
        self.assertGreaterEqual(len(prs.slides), j['report']['slides'] - 1)
        self.assertEqual(len(prs.slides), j['report']['slides'])
        # ảnh gốc đã xoá sau khi xuất
        self.assertFalse(os.path.isdir(os.path.join(self.S.JOBS, jid, 'photos')))
        st, d = self.req('GET', '/api/jobs', self.staff)
        row = next(x for x in d['jobs'] if x['id'] == jid)
        self.assertGreater(row['expires_ts'], row['finished_ts'])
        # Mở lại cài đặt của lần xuất
        st, d = self.req('GET', f'/api/jobs/{jid}/settings', self.staff)
        self.assertEqual(st, 200, d)
        self.assertEqual(d['preset']['opts']['n_per_slide'], 2)
        self.assertEqual(d['title'], 'Báo cáo thử W1')
        # Tải nhiều kết quả thành .zip
        import zipfile
        st, b = self.req('GET', f'/api/jobs/zip?ids={jid}', self.staff, raw=True)
        self.assertEqual(st, 200)
        names = zipfile.ZipFile(io.BytesIO(b)).namelist()
        self.assertIn('Báo cáo thử W1/Báo cáo thử W1.pptx', names)
        st, _ = self.req('GET', f'/api/jobs/zip?ids={jid}', self.admin.replace(self.admin, 'sai'))
        self.assertEqual(st, 401)
        st, b = self.req('GET', '/static/ga-symbol.png', raw=True)
        self.assertEqual((st, b[:4]), (200, b'\x89PNG'))
        st, d = self.req('DELETE', f'/api/jobs/{jid}', self.staff)
        self.assertEqual(st, 200)

    def test_analyze_groups_from_names(self):
        """Trang biên tập: gửi tên ảnh (chưa tải lên) → nhóm theo mã, khớp Excel, avatar."""
        hdr = {'X-File-Name': urllib.parse.quote('An.xlsx')}
        st, d = self.req('PUT', '/api/excel?target=all', self.admin,
                         make_xlsx(['ANA01', 'ANA02']), hdr)
        self.assertEqual(st, 200, d)
        self.req('PUT', '/api/files/avatars/Building/ANA01A.jpg', self.admin, make_jpeg(40, 40))
        files = ['f1/W1/ANA01 (1).jpg', 'f1/W1/ANA01 (2).jpg', 'f1/W1/ANA02 (1).jpg',
                 'f1/W1/LAZZ9 (1).jpg', 'f1/W1/ghi-chu.txt']
        st, d = self.req('POST', '/api/analyze', self.staff,
                         {'files': files, 'preset_id': 'builtin:Báo cáo'})
        self.assertEqual(st, 200, d)
        by = {g['code']: g for g in d['groups']}
        self.assertEqual(set(by), {'ANA01', 'ANA02', 'LAZZ9'})
        self.assertEqual(by['ANA01']['files'], ['f1/W1/ANA01 (1).jpg', 'f1/W1/ANA01 (2).jpg'])
        self.assertEqual(by['ANA01']['kind'], 'exact')
        self.assertEqual(by['ANA01']['avatar'], 'Building/ANA01A.jpg')
        self.assertEqual(by['ANA01']['name'], 'Điểm ANA01')
        self.assertTrue(any(r[0] == 'Address' for r in by['ANA01']['info_rows']))
        self.assertEqual(by['LAZZ9']['kind'], 'none')
        self.assertEqual(d['photos'], 4)
        # SALESKIT trả bảng riêng
        st, d = self.req('POST', '/api/analyze', self.staff,
                         {'files': files, 'preset_id': 'builtin:SALESKIT'})
        g = next(x for x in d['groups'] if x['code'] == 'ANA01')
        self.assertIn('specs', g['info_table'])

    def test_ui_served(self):
        st, b = self.req('GET', '/', raw=True)
        self.assertEqual(st, 200)
        self.assertIn(b'AutoPPTX', b)


if __name__ == '__main__':
    unittest.main()
