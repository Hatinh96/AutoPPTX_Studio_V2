import unittest

from autopptx2.effects import _random_dt, _dt_from, _parse_hm


class TestRandomDatetime(unittest.TestCase):
    def test_parse_hm(self):
        self.assertEqual(_parse_hm('09:30'), (9, 30))
        self.assertEqual(_parse_hm('bad', 8, 0), (8, 0))

    def test_random_time_in_range(self):
        fx = {
            'random_date': False,
            'fix_date': True,
            'fix_date_val': '15/08/2026',
            'random_time_from': '10:00',
            'random_time_to': '12:00',
        }
        dt = _random_dt('/photos/site_a/img001.jpg', fx)
        self.assertEqual(dt.strftime('%d/%m/%Y'), '15/08/2026')
        mins = dt.hour * 60 + dt.minute
        self.assertGreaterEqual(mins, 10 * 60)
        self.assertLessEqual(mins, 12 * 60)

    def test_same_path_same_time(self):
        fx = {
            'random_date': True,
            'random_date_from': '01/08/2026',
            'random_date_to': '10/08/2026',
            'random_time_from': '08:00',
            'random_time_to': '18:00',
        }
        p = '/x/photo_zalo.jpg'
        self.assertEqual(_random_dt(p, fx), _random_dt(p, fx))

    def test_dt_from_uses_random_without_exif(self):
        fx = {
            'use_exif': True,
            'random_no_exif': True,
            'random_date': False,
            'fix_date': True,
            'fix_date_val': '20/08/2026',
            'random_time_from': '09:00',
            'random_time_to': '11:00',
            'fix_date': True,
            'fix_time': False,
        }
        dt = _dt_from('/no/exif/zalo_photo.jpg', fx)
        self.assertEqual(dt.strftime('%d/%m/%Y'), '20/08/2026')
        self.assertGreaterEqual(dt.hour, 9)
        self.assertLessEqual(dt.hour, 11)


if __name__ == '__main__':
    unittest.main()
