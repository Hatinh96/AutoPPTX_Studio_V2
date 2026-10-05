"""Bộ nhớ đệm khớp mã / avatar phải cho kết quả y như tính lại từ đầu."""
import os
import tempfile
import unittest

from autopptx2.datasource import AvatarIndex, match_row


def _rows(*codes):
    return {c: {'Code_RP': c, 'Name': c.title(), '_SiteStatus': 'on'} for c in codes}


class TestMatchRowCache(unittest.TestCase):
    def test_repeat_calls_same_result(self):
        by_code = _rows('HLVINCOMBATRIEU', 'HLNGUYENTRAI', 'PLHOANGDIEU')
        first = match_row('HLVINCOMBATRIE', by_code)
        for _ in range(3):
            self.assertEqual(match_row('HLVINCOMBATRIE', by_code), first)
        self.assertEqual(first[2], 'fuzzy')
        self.assertEqual(first[1], 'HLVINCOMBATRIEU')

    def test_new_row_invalidates_cached_miss(self):
        by_code = _rows('HLNGUYENTRAI')
        self.assertEqual(match_row('PLHOANGDIEX', by_code)[2], 'none')
        by_code['PLHOANGDIEU'] = {'Code_RP': 'PLHOANGDIEU', '_SiteStatus': 'on'}
        row, code, kind = match_row('PLHOANGDIEX', by_code)
        self.assertEqual((code, kind), ('PLHOANGDIEU', 'fuzzy'))

    def test_separate_tables_do_not_leak(self):
        a = _rows('HLVINCOMBATRIEU')
        b = _rows('PLHOANGDIEU')
        self.assertEqual(match_row('HLVINCOMBATRIE', a)[1], 'HLVINCOMBATRIEU')
        self.assertEqual(match_row('HLVINCOMBATRIE', b)[2], 'none')

    def test_none_result_is_fresh_dict(self):
        by_code = _rows('HLNGUYENTRAI')
        r1, _, _ = match_row('ZZZZZZZZ', by_code)
        r1['x'] = 1
        r2, _, _ = match_row('ZZZZZZZZ', by_code)
        self.assertEqual(r2, {})


class TestAvatarMemo(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        for name in ('HLNGUYENTRAIA.jpg', 'PLHOANGDIEUA.jpg'):
            with open(os.path.join(self.tmp.name, name), 'wb') as f:
                f.write(b'x')
        self.idx = AvatarIndex()
        self.idx.set_folder(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_fuzzy_result_stable(self):
        p = self.idx.find('PLHOANGDIEX')
        self.assertTrue(p and p.endswith('PLHOANGDIEUA.jpg'))
        self.assertEqual(self.idx.find('PLHOANGDIEX'), p)

    def test_set_folder_resets_memo(self):
        self.assertIsNone(self.idx.find('BRANDNEWSTORE'))
        with open(os.path.join(self.tmp.name, 'BRANDNEWSTORE.jpg'), 'wb') as f:
            f.write(b'x')
        self.idx.set_folder(self.tmp.name)
        self.assertTrue(self.idx.find('BRANDNEWSTORE'))

    def test_set_aliases_resets_memo(self):
        self.assertIsNone(self.idx.find('OLDCODE'))
        self.idx.set_aliases({'OLDCODE': 'HLNGUYENTRAI'})
        self.assertTrue(self.idx.find('OLDCODE').endswith('HLNGUYENTRAIA.jpg'))


if __name__ == '__main__':
    unittest.main()
