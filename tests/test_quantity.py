"""Quantity trên slide = tổng màn, không phải cột Quantity (số phòng)."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from autopptx2.datasource import (
    _quantity_text, build_info_rows, build_info_table, screen_qty,
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
        self.assertEqual(sorted(qtys), ['2', '6'])
        self.assertNotIn('60', qtys)

    def test_saleskit_coffee_uses_quantity_column(self):
        row = {
            'Name': 'Coffee A', 'Location': 'Quầy',
            'Quantity': 3, 'LCD': 0, 'DP': 0, 'GP': 0,
        }
        table = build_info_table(row, [row])
        self.assertEqual([s['qty'] for s in table['specs']], ['3'])


if __name__ == '__main__':
    unittest.main()
