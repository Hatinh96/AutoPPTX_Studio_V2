import unittest

from autopptx2 import geometry as G


class TestSlideBatches(unittest.TestCase):
    def test_pad_blank_by_screen_count(self):
        paths = ['a.jpg', 'b.jpg', 'c.jpg', 'd.jpg']
        batches = G.build_slide_batches(paths, 4, screen_count=28, pad_blank=True)
        self.assertEqual(len(batches), 7)
        self.assertEqual(len(batches[0]), 4)
        self.assertEqual(batches[0], paths)
        self.assertEqual(batches[1][0], None)

    def test_pad_blank_keeps_extra_photos(self):
        paths = [f'{i}.jpg' for i in range(10)]
        batches = G.build_slide_batches(paths, 4, screen_count=4, pad_blank=True)
        flat = [p for b in batches for p in b if p]
        self.assertEqual(flat, paths)
        self.assertGreaterEqual(sum(len(b) for b in batches), 10)

    def test_no_pad_photo_only(self):
        paths = ['a.jpg', 'b.jpg']
        batches = G.build_slide_batches(paths, 4, screen_count=28, pad_blank=False)
        self.assertEqual(len(batches), 1)
        self.assertEqual(len(batches[0]), 2)


if __name__ == '__main__':
    unittest.main()
