"""Chia file PPTX: 0 = một file, số dương = cắt theo N slide."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autopptx2 import geometry as G
from autopptx2.exporter import estimate


class TestSlidesPerFile(unittest.TestCase):
    def test_parse_one_file(self):
        self.assertEqual(G.parse_slides_per_file(0), 0)
        self.assertEqual(G.parse_slides_per_file('Một file'), 0)
        self.assertEqual(G.parse_slides_per_file('gộp'), 0)
        self.assertEqual(G.slides_per_file_label(0), 'Một file')

    def test_parse_number(self):
        self.assertEqual(G.parse_slides_per_file(200), 200)
        self.assertEqual(G.parse_slides_per_file('80'), 80)
        self.assertEqual(G.parse_slides_per_file('abc', default=150), 150)
        self.assertEqual(G.file_part_limit(0), 0)
        self.assertEqual(G.file_part_limit(200), 200)

    def test_part_count(self):
        self.assertEqual(G.file_part_count(0, 200), 1)
        self.assertEqual(G.file_part_count(400, 0), 1)
        self.assertEqual(G.file_part_count(400, 200), 2)
        self.assertEqual(G.file_part_count(401, 200), 3)
        self.assertEqual(G.file_part_count(80, 200), 1)

    def test_estimate_one_file(self):
        groups = {f'C{i}': ['a.jpg'] for i in range(400)}
        ng, ns, np_ = estimate(groups, 1, slides_per_file=0)
        self.assertEqual(ng, 400)
        self.assertEqual(ns, 400)
        self.assertEqual(np_, 1)

    def test_estimate_split_200(self):
        groups = {f'C{i}': ['a.jpg'] for i in range(400)}
        _ng, ns, np_ = estimate(groups, 1, slides_per_file=200)
        self.assertEqual(ns, 400)
        self.assertEqual(np_, 2)

    def test_estimate_custom_100(self):
        groups = {f'C{i}': ['a.jpg'] for i in range(250)}
        _ng, _ns, np_ = estimate(groups, 1, slides_per_file=100)
        self.assertEqual(np_, 3)


if __name__ == '__main__':
    unittest.main()
