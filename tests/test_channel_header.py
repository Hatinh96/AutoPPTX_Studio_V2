"""Cột kênh: nhận Channel / Chanel / Kênh để lọc từng kênh."""
import os
import sys
import tempfile
import unittest

from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autopptx2.datasource import ExcelSource, _norm_header


class TestChannelHeader(unittest.TestCase):
    def test_norm_kenh_and_typo(self):
        self.assertEqual(_norm_header('Channel'), 'channel')
        self.assertEqual(_norm_header('Chanel'), 'chanel')
        self.assertEqual(_norm_header('Kênh'), 'kenh')
        self.assertEqual(_norm_header('Kênh quảng cáo'), 'kenhquangcao')

    def test_load_chanel_typo(self):
        wb = Workbook()
        ws = wb.active
        ws.append(['No.', 'BD Code', 'Chanel', 'Report Code', 'Name'])
        ws.append([1, 'X', 'CF', 'HIGHLANDSQ1', 'Highlands'])
        ws.append([2, 'Y', 'MT', 'KOICFEQ1', 'KOI'])
        fd, path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        try:
            wb.save(path)
            src = ExcelSource()
            self.assertTrue(src.load(path), src.error)
            self.assertEqual(src.channels(), ['CF', 'MT'])
            self.assertEqual(src.rows[0]['Channel'], 'CF')
        finally:
            os.remove(path)

    def test_cloud_coffee_cache_if_present(self):
        path = os.path.join(
            os.path.expanduser('~'), '.autopptx_studio_v2', 'excel_cache',
            '785e6789-db65-4f34-bd31-a69140fe0aad.xlsx')
        if not os.path.isfile(path):
            self.skipTest('no local cloud excel cache')
        src = ExcelSource()
        self.assertTrue(src.load(path), src.error)
        self.assertGreater(len(src.rows), 0)
        self.assertIn('CF', src.channels())
        self.assertGreaterEqual(len(src.channels()), 1)


if __name__ == '__main__':
    unittest.main()
