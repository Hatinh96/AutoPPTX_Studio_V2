"""Avatar: khớp đúng Code_RP, gần giống (difflib 0.87), không lấy ảnh đầu thư mục."""
import os
import sys
import tempfile
import unittest

from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from autopptx2.datasource import AvatarIndex


def _touch_jpg(folder, name):
    path = os.path.join(folder, name)
    Image.new('RGB', (8, 8), (10, 20, 30)).save(path, 'JPEG')
    return path


class AvatarIndexTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.folder = self._tmp.name
        self.idx = AvatarIndex()

    def tearDown(self):
        self._tmp.cleanup()

    def test_exact_match(self):
        want = _touch_jpg(self.folder, 'HLHOANGDIEU2FC.jpg')
        _touch_jpg(self.folder, 'HLVINCOMLEVANVIET.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('HLHOANGDIEU2FC'), want)
        self.assertEqual(self.idx.find('hlhoangdieu2fc'), want)

    def test_suffix_ava_and_a(self):
        want = _touch_jpg(self.folder, 'KFCVOVANNGAN_Ava.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('KFCVOVANNGAN'), want)
        aa = _touch_jpg(self.folder, 'LOTTERIALEVANVIETAA.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('LOTTERIALEVANVIET'), aa)

    def test_alnum_normalize(self):
        want = _touch_jpg(self.folder, 'HL-VINCOM-LE-VAN-VIET.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('HLVINCOMLEVANVIET'), want)

    def test_block_alias_uses_root_not_neighbor(self):
        root = _touch_jpg(self.folder, 'TOPAZHOME2B.jpg')
        b2 = _touch_jpg(self.folder, 'TOPAZHOME2BBLOCKB2.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('TOPAZHOME2BBLOCKB1'), root)
        self.assertEqual(self.idx.find('TOPAZHOME2BBLOCKB2'), b2)
        self.assertEqual(self.idx.find('TOPAZHOME2B'), root)

    def test_fuzzy_neighbor_block(self):
        """Thiếu B1, chỉ có B2 → dùng ảnh B2 (cutoff 0.87)."""
        b2 = _touch_jpg(self.folder, 'TOPAZHOME2BBLOCKB2.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('TOPAZHOME2BBLOCKB2'), b2)
        self.assertEqual(self.idx.find('TOPAZHOME2BBLOCKB1'), b2)
        self.assertEqual(self.idx.find('TOPAZHOME2BBLOCKB3'), b2)

    def test_missing_is_none_not_first_file(self):
        first = _touch_jpg(self.folder, 'AAA_FIRST.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('AAA_FIRST'), first)
        self.assertIsNone(self.idx.find('MA_KHONG_CO_AVATAR'))
        self.assertIsNone(self.idx.resolve('MA_KHONG_CO_AVATAR', 'CUNG_KHONG_CO'))
        self.assertFalse(hasattr(self.idx, 'first'))

    def test_code_ending_with_a_keeps_own_file(self):
        own = _touch_jpg(self.folder, 'SUNVIEWRUBYA.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.find('SUNVIEWRUBYA'), own)

    def test_explicit_alias_table(self):
        dest = _touch_jpg(self.folder, 'HLDANGVANBI.jpg')
        self.idx.set_folder(self.folder)
        self.idx.set_aliases({'HL-DANG-VAN-BI-OLD': 'HLDANGVANBI'})
        self.assertEqual(self.idx.find('HL-DANG-VAN-BI-OLD'), dest)
        self.assertIsNone(self.idx.find('HL-KHONG-CO'))

    def test_resolve_prefers_image_code(self):
        a = _touch_jpg(self.folder, 'CODEA.jpg')
        b = _touch_jpg(self.folder, 'CODEB.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.resolve('CODEA', 'CODEB'), a)
        self.assertEqual(self.idx.resolve(None, 'CODEB'), b)
        self.assertIsNone(self.idx.resolve('', None))

    def test_fuzzy_via_resolve_and_ava_suffix(self):
        want = _touch_jpg(self.folder, 'TOPAZHOME2BBLOCKB2_Ava.jpg')
        self.idx.set_folder(self.folder)
        self.assertEqual(self.idx.resolve('TOPAZHOME2BBLOCKB1'), want)
        self.assertEqual(self.idx.find('TOPAZHOME2BBLOCKB1'), want)


if __name__ == '__main__':
    unittest.main()
