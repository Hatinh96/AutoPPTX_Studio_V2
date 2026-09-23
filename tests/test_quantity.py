"""Quantity trên slide = tổng màn, không phải cột Quantity (số phòng)."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from autopptx2.datasource import (
    _quantity_text, build_info_rows, build_info_table, screen_qty,
    format_district_display, wrap_info_cell, info_row_weights,
)


class QuantityScreensTests(unittest.TestCase):
    def test_hotel_ignores_room_quantity(self):
        row = {'Quantity': 60, 'LCD': 2, 'DP': 6, 'GP': 0}
        self.assertEqual(screen_qty(row), 8)
        self.assertEqual(_quantity_text(row), '8 (2 LCD, 6 DP)')
        self.assertEqual(build_info_rows(row)[2],
                         ('Quantity', '8 (2 LCD, 6 DP)', False))

    def test_includes_ds_like_v1(self):
        row = {'Quantity': 40, 'LCD': 1, 'DP': 2, 'GP': 3, 'DS': 4}
        self.assertEqual(_quantity_text(row), '10 (1 LCD, 2 DP, 4 DS, 3 GP)')

    def test_includes_led_dps_dpf(self):
        row = {'Quantity': 99, 'LCD': 1, 'LED': 2, 'DPS': 3, 'DPF': 4}
        self.assertEqual(_quantity_text(row), '10 (1 LCD, 3 DPS, 4 DPF, 2 LED)')

    def test_nha_trang_dp_ds_breakdown(self):
        row = {'Quantity': 2, 'LCD': 0, 'DP': 1, 'DS': 1, 'GP': 0}
        self.assertEqual(_quantity_text(row), '2 (1 DP, 1 DS)')

    def test_coffee_falls_back_to_quantity_column(self):
        row = {'Quantity': 3, 'LCD': 0, 'DP': 0, 'GP': 0}
        self.assertEqual(screen_qty(row), 0)
        self.assertEqual(_quantity_text(row), '3')

    def test_blank_falls_back_to_file_count(self):
        self.assertEqual(_quantity_text({}, n_files=5), '5')
        self.assertEqual(_quantity_text({}), '—')

    def test_saleskit_qty_is_screens_not_rooms(self):
        row = {
            'Name': 'Hotel A', 'Location': 'Lễ tân',
            'Quantity': 60, 'LCD': 2, 'DP': 6, 'GP': 0, 'Size': '32"',
        }
        table = build_info_table(row, [row])
        qtys = [s['qty'] for s in table['specs']]
        self.assertEqual(qtys, ['8 (2 LCD, 6 DP)'])
        self.assertEqual([s['size'] for s in table['specs']], ['32"'])
        self.assertNotIn('60', qtys)

    def test_saleskit_groups_dp_family_and_keeps_gp_separate(self):
        row = {
            'Name': 'University A', 'Location': 'Khu A',
            'Quantity': 27, 'LCD': 3, 'DP': 16,
            'DPS': 6, 'DPF': 2, 'GP': 4, 'Size': '32"',
        }
        table = build_info_table(row, [row])
        self.assertEqual(
            [(spec['form'], spec['qty']) for spec in table['specs']],
            [
                ('LCD, DP, DPS, DPF', '27 (3 LCD, 16 DP, 6 DPS, 2 DPF)'),
                ('GP', '4'),
            ],
        )
        self.assertEqual([spec['size'] for spec in table['specs']], ['32"', ''])

    def test_saleskit_uses_university_gp_placement_details(self):
        row = {
            'Name': 'University GP', 'Location': 'Khu B',
            'Channel': 'University (DP.LCD)',
            'DP': 4, 'LCD': 2, 'GP': 16,
            'GP_Inside': 2, 'GP_Ground': 2,
            'GP_Parking': 0, 'GP_Study': 12,
            'Size': '32" CVĐ',
        }
        table = build_info_table(row, [row])
        self.assertEqual(
            [(spec['form'], spec['qty']) for spec in table['specs']],
            [
                ('LCD, DP', '6 (2 LCD, 4 DP)'),
                ('GP', '16'),
            ],
        )
        self.assertEqual(
            [spec['size'] for spec in table['specs']],
            ['32" CVĐ', ''],
        )

    def test_saleskit_coffee_uses_quantity_column(self):
        row = {
            'Name': 'Coffee A', 'Location': 'Quầy',
            'Quantity': 3, 'LCD': 0, 'DP': 0, 'GP': 0,
        }
        table = build_info_table(row, [row])
        self.assertEqual([s['qty'] for s in table['specs']], ['3'])

    def test_saleskit_building_shows_gp_forms(self):
        row = {
            'Name': 'Building A', 'District': 'Thủ Đức',
            'Channel': 'Building', 'Quantity': 3,
            'GP_Inside': 3, 'GP_Ground': 8,
            'GP_Facilities': 0, 'GP_Parking': 2,
            'Size': '27" CVĐ',
            'GP_Size': '120 x 180 cm',
        }
        table = build_info_table(row, [row])
        self.assertEqual(
            [(spec['form'], spec['qty']) for spec in table['specs']],
            [('GP', '13')],
        )
        self.assertEqual(
            [spec['size'] for spec in table['specs']],
            ['120 x 180 cm'],
        )
        self.assertEqual(screen_qty(row), 13)

    def test_qty_show_hides_gp_from_screen_count(self):
        row = {'LCD': 2, 'DP': 6, 'GP': 10}
        self.assertEqual(screen_qty(row), 18)
        self.assertEqual(screen_qty(row, {'GP': False}), 8)


class InfoTableDisplayTests(unittest.TestCase):
    def test_district_strips_quan_prefix(self):
        self.assertEqual(format_district_display('Quận 9'), '9')
        self.assertEqual(format_district_display('Q. 1'), '1')
        self.assertEqual(build_info_rows({'District': 'Quận 9'})[1],
                         ('District', '9', False))

    def test_address_wraps_long_text(self):
        addr = '50 Lê Văn Việt, Khu công nghệ cao, TP. Thủ Đức'
        lines = wrap_info_cell(addr, max_chars=18)
        self.assertGreater(len(lines), 1)
        weights = info_row_weights([('Address', addr, False)], 3.0)
        self.assertGreater(weights[0], 1.0)


if __name__ == '__main__':
    unittest.main()
