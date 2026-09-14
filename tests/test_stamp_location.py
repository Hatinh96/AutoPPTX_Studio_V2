import unittest
from unittest.mock import patch
from PIL import Image

from autopptx2.datasource import excel_stamp_location, location_fx_fields
from autopptx2.effects import (
    DEFAULT_FX, location_stamp_lines, render_timestamp, resolve_location,
)


class TestStampLocationToggle(unittest.TestCase):
    def _blank(self, w=400, h=300):
        return Image.new('RGBA', (w, h), (40, 40, 40, 255))

    @patch('autopptx2.effects.resolve_location', return_value='Yersin\nPhường Nha Trang\nKhánh Hòa')
    @patch('autopptx2.effects._dt_from')
    def test_date_only_without_location(self, mock_dt, mock_loc):
        mock_dt.return_value = __import__('datetime').datetime(2026, 8, 2, 10, 52)
        fx = dict(DEFAULT_FX)
        fx.update({'use_timestamp': True, 'stamp_location': False, 'show_gps': False})
        im = render_timestamp(self._blank(), '/fake.jpg', fx, allow_net=False)
        txt = im.tobytes()  # rendered — inspect draw via mock calls on loc
        mock_loc.assert_not_called()

    @patch('autopptx2.effects.resolve_location', return_value='Quận 1')
    @patch('autopptx2.effects._dt_from')
    def test_location_when_enabled(self, mock_dt, mock_loc):
        mock_dt.return_value = __import__('datetime').datetime(2026, 8, 2, 10, 52)
        fx = dict(DEFAULT_FX)
        fx.update({'use_timestamp': True, 'stamp_location': True, 'show_gps': False})
        render_timestamp(self._blank(), '/fake.jpg', fx, allow_net=False)
        mock_loc.assert_called_once()


class TestExcelStampAddress(unittest.TestCase):
    def test_address_district_vietnam_not_shop_name(self):
        row = {
            'Name': 'Highlands Coffee Calmette',
            'Address': '53 Võ Văn Ngân',
            'District': 'Thủ Đức',
        }
        text = excel_stamp_location(row)
        self.assertEqual(text, '53 Võ Văn Ngân\nThủ Đức\nViệt Nam')
        self.assertNotIn('Highlands', text)
        self.assertEqual(location_fx_fields(row)['location_excel'], text)

    def test_numeric_district_gets_quan_prefix(self):
        row = {'Address': '161-163 Calmette', 'District': '1', 'Name': 'HL'}
        self.assertEqual(
            excel_stamp_location(row),
            '161-163 Calmette\nQuận 1\nViệt Nam')

    def test_stamp_lines_keep_vietnam(self):
        self.assertEqual(
            location_stamp_lines('53 Võ Văn Ngân\nThủ Đức\nViệt Nam'),
            ['53 Võ Văn Ngân', 'Thủ Đức', 'Việt Nam'])

    def test_excel_mode_uses_address_block(self):
        fx = dict(DEFAULT_FX)
        fx.update({
            'location_mode': 'excel',
            'location_excel': '53 Võ Văn Ngân\nThủ Đức\nViệt Nam',
        })
        self.assertEqual(
            resolve_location('/x.jpg', fx, allow_net=False),
            '53 Võ Văn Ngân\nThủ Đức\nViệt Nam')


if __name__ == '__main__':
    unittest.main()

