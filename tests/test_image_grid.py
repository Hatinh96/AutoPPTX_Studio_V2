"""Lưới ảnh: số ô = số ảnh thật, Tự = tối đa 4."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autopptx2 import geometry as G


class TestImageGrid(unittest.TestCase):
    def test_per_slide_max_auto(self):
        self.assertEqual(G.per_slide_max(0), 4)
        self.assertEqual(G.per_slide_max('Tự'), 4)
        self.assertEqual(G.per_slide_max('auto'), 4)
        self.assertEqual(G.per_slide_max(2), 2)
        self.assertEqual(G.per_slide_max(4), 4)

    def test_grid_dims_follow_count(self):
        self.assertEqual(G.grid_dims(1), (1, 1))
        self.assertEqual(G.grid_dims(2), (2, 1))
        self.assertEqual(G.grid_dims(3), (2, 2))
        self.assertEqual(G.grid_dims(4), (2, 2))
        self.assertEqual(G.grid_dims(6), (3, 2))

    def test_one_photo_fills_area(self):
        cells = G.grid_cells((0, 0, 12.0, 4.0), 1, None, 0.1)
        self.assertEqual(len(cells), 1)
        _x, _y, w, h = cells[0]
        self.assertGreater(w, 10.0)
        self.assertGreater(h, 3.0)

    def test_no_empty_pad(self):
        paths = ['a.jpg']
        self.assertEqual(G.slide_images(paths, 2), ['a.jpg'])
        self.assertEqual(len(G.grid_cells((0, 0, 10, 4), 1, 4 / 3, 0.1)), 1)

    def test_two_side_by_side(self):
        cells = G.grid_cells((0, 0, 12.0, 4.0), 2, None, 0.1)
        self.assertEqual(len(cells), 2)
        self.assertLess(cells[0][0], cells[1][0])


if __name__ == '__main__':
    unittest.main()
