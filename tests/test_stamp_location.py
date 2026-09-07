import unittest
from unittest.mock import patch
from PIL import Image

from autopptx2.effects import render_timestamp, DEFAULT_FX


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


if __name__ == '__main__':
    unittest.main()
