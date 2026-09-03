"""Thứ tự tỉnh/thành Bắc → Nam cho sales kit BD."""
import os
import sys
import unittest
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autopptx2.datasource import (
    norm_city, city_sort_key, sort_codes_by_city, order_groups_by_city,
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


if __name__ == '__main__':
    unittest.main()
