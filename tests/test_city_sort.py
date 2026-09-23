"""Thứ tự tỉnh/thành Bắc → Nam cho Sales và BD."""
import os
import sys
import unittest
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autopptx2.datasource import (
    norm_city, city_sort_key, sort_codes_by_city, order_groups_by_city,
    sanitize_place_fields, _address_line, place_is_city, is_province_name,
    scope_by_channel,
)


class TestCitySort(unittest.TestCase):
    def test_norm_aliases(self):
        self.assertEqual(norm_city('Hồ Chí Minh'), 'ho chi minh')
        self.assertEqual(norm_city('Hồ chí Minh'), 'ho chi minh')
        self.assertEqual(norm_city('TP.HCM'), 'ho chi minh')
        self.assertEqual(norm_city('Binh Duong'), 'binh duong')
        self.assertEqual(norm_city('Bình Dương'), 'binh duong')
        self.assertEqual(norm_city('Thành phố Huế'), 'hue')
        self.assertEqual(norm_city('Hà Nội'), 'ha noi')

    def test_north_before_south(self):
        hn = city_sort_key('Hà Nội')
        dn = city_sort_key('Đà Nẵng')
        hcm = city_sort_key('Hồ Chí Minh')
        ct = city_sort_key('Cần Thơ')
        empty = city_sort_key('')
        self.assertLess(hn, dn)
        self.assertLess(dn, hcm)
        self.assertLess(hcm, ct)
        self.assertLess(hcm, empty)

    def test_sort_codes(self):
        by_code = {
            'A': {'City': 'Hồ Chí Minh', 'Name': 'Z', 'District': ''},
            'B': {'City': 'Hà Nội', 'Name': 'Y', 'District': ''},
            'C': {'City': 'Đà Nẵng', 'Name': 'X', 'District': ''},
        }
        self.assertEqual(sort_codes_by_city(['A', 'B', 'C'], by_code),
                         ['B', 'C', 'A'])

    def test_order_groups(self):
        groups = OrderedDict([
            ('HCM1', ['a.jpg']),
            ('HN1', ['b.jpg']),
        ])
        by_code = {
            'HCM1': {'City': 'Hồ Chí Minh', 'Name': 'S', 'District': ''},
            'HN1': {'City': 'Hà Nội', 'Name': 'N', 'District': ''},
        }
        out = order_groups_by_city(groups, by_code)
        self.assertEqual(list(out), ['HN1', 'HCM1'])

    def test_city_in_province_stays_as_district(self):
        """'Vũng Tàu' là TP thuộc tỉnh — District hợp lệ, không được xoá."""
        rec = {'City': 'Bà Rịa - Vũng Tàu', 'District': 'Vũng Tàu',
               'Ward': '', 'Address': '01 Trần Phú'}
        sanitize_place_fields(rec)
        self.assertEqual(rec['District'], 'Vũng Tàu')
        line = _address_line(rec)
        self.assertNotIn('Phường Vũng Tàu', line)
        self.assertIn('Vũng Tàu', line)
        self.assertIn('01 Trần Phú', line)
        self.assertFalse(is_province_name('Vũng Tàu'))
        self.assertTrue(place_is_city('Vũng Tàu'))

    def test_province_in_district_is_cleared(self):
        rec = {'City': 'Đà Nẵng', 'District': 'Đà Nẵng', 'Ward': '',
               'Address': '12 Bạch Đằng'}
        sanitize_place_fields(rec)
        self.assertEqual(rec['District'], '')
        self.assertNotIn('Phường', _address_line(rec))

    def test_university_ward_promoted_to_district(self):
        rec = {'City': 'Hà Nội', 'District': 'Hà Nội', 'Ward': 'Bắc Từ Liêm'}
        sanitize_place_fields(rec)
        self.assertEqual(rec['District'], 'Bắc Từ Liêm')

    def test_empty_city_takes_province_from_district(self):
        rec = {'City': '', 'District': 'Hồ Chí Minh', 'Ward': ''}
        sanitize_place_fields(rec)
        self.assertEqual(rec['City'], 'Hồ Chí Minh')
        self.assertEqual(rec['District'], '')

    def test_scope_by_channel_limits_qa_scope(self):
        by_code = {
            'UNI1': {'Channel': 'University'},
            'CF1': {'Channel': 'Coffee Shop & Milk Tea'},
        }
        out = scope_by_channel(by_code, 'University', 'Tất cả kênh')
        self.assertEqual(list(out), ['UNI1'])
        self.assertEqual(list(scope_by_channel(by_code, 'Tất cả kênh',
                                               'Tất cả kênh')), ['UNI1', 'CF1'])


if __name__ == '__main__':
    unittest.main()
