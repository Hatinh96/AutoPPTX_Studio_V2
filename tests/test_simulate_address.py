import unittest
from unittest.mock import patch

from autopptx2.effects import (
    DEFAULT_FX, parse_address_list, resolve_location, simulated_address,
)


class TestSimulateAddress(unittest.TestCase):
    def test_parse_skips_blank_and_dupes(self):
        text = "12 Nguyễn Huệ, Q.1\n\n• 12 Nguyễn Huệ, Q.1\n- 99 Lê Lợi, Q.1\n"
        self.assertEqual(
            parse_address_list(text),
            ['12 Nguyễn Huệ, Q.1', '99 Lê Lợi, Q.1'])

    def test_single_list_line_wins(self):
        fx = {'location_list': '12 Nguyễn Huệ, Q.1', 'location_address': 'Khác'}
        self.assertEqual(simulated_address('/a.jpg', fx), '12 Nguyễn Huệ, Q.1')

    def test_excel_address_when_list_has_many(self):
        fx = {
            'location_list': 'A\nB\nC',
            'location_address': 'Tòa nhà Bitexco, Q.1, TP.HCM',
        }
        self.assertEqual(
            simulated_address('/site/a.jpg', fx),
            'Tòa nhà Bitexco, Q.1, TP.HCM')

    def test_list_pick_is_stable(self):
        fx = {'location_list': 'Một\nHai\nBa'}
        a = simulated_address('/photos/zalo_1.jpg', fx)
        self.assertEqual(a, simulated_address('/photos/zalo_1.jpg', fx))
        self.assertIn(a, ('Một', 'Hai', 'Ba'))

    @patch('autopptx2.effects.gps_cached', return_value=None)
    @patch('autopptx2.effects.location_from_exif', return_value='')
    def test_gps_mode_uses_list_when_no_exif(self, _exif, _gps):
        fx = dict(DEFAULT_FX)
        fx.update({
            'location_mode': 'gps',
            'simulate_no_gps': True,
            'location_list': '45 Pasteur, Q.1, TP.HCM',
        })
        self.assertEqual(
            resolve_location('/no-gps.jpg', fx, allow_net=False),
            '45 Pasteur, Q.1, TP.HCM')

    @patch('autopptx2.effects.gps_cached', return_value=None)
    @patch('autopptx2.effects.location_from_exif', return_value='')
    def test_off_keeps_blank_when_no_gps(self, _exif, _gps):
        fx = dict(DEFAULT_FX)
        fx.update({
            'location_mode': 'gps',
            'simulate_no_gps': False,
            'location_list': '45 Pasteur, Q.1',
        })
        self.assertEqual(resolve_location('/no-gps.jpg', fx, allow_net=False), '')

    @patch('autopptx2.effects.gps_cached', return_value=(10.77, 106.70))
    @patch('autopptx2.effects.location_from_exif', return_value='Đường Lê Lợi')
    def test_real_gps_not_replaced(self, _exif, _gps):
        fx = dict(DEFAULT_FX)
        fx.update({
            'location_mode': 'gps',
            'simulate_no_gps': True,
            'simulate_all': False,
            'location_list': 'Địa chỉ giả',
        })
        self.assertEqual(
            resolve_location('/has-gps.jpg', fx, allow_net=False),
            'Đường Lê Lợi')

    @patch('autopptx2.effects.gps_cached', return_value=(10.77, 106.70))
    @patch('autopptx2.effects.location_from_exif', return_value='Đường Lê Lợi')
    def test_simulate_all_overrides_real_gps(self, _exif, _gps):
        fx = dict(DEFAULT_FX)
        fx.update({
            'location_mode': 'gps',
            'simulate_no_gps': False,
            'simulate_all': True,
            'location_list': '45 Pasteur, Q.1, TP.HCM',
        })
        self.assertEqual(
            resolve_location('/has-gps.jpg', fx, allow_net=False),
            '45 Pasteur, Q.1, TP.HCM')


if __name__ == '__main__':
    unittest.main()
