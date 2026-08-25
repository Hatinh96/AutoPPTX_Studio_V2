import os
import unittest
from unittest import mock

from autopptx2 import fonts


class FontPlatformTests(unittest.TestCase):
    def test_macos_font_dirs_include_user_library_and_system_fonts(self):
        existing = {
            os.path.abspath(os.path.expanduser('~/Library/Fonts')),
            os.path.abspath('/Library/Fonts'),
            os.path.abspath('/System/Library/Fonts/Supplemental'),
            os.path.abspath('/System/Library/Fonts'),
        }
        with mock.patch.object(fonts.sys, 'platform', 'darwin'), \
                mock.patch.object(fonts.os.path, 'isdir',
                                  side_effect=lambda p: p in existing):
            found = fonts.font_dirs()

        self.assertIn(os.path.abspath(os.path.expanduser('~/Library/Fonts')), found)
        self.assertIn(os.path.abspath('/System/Library/Fonts'), found)
        self.assertIn(
            os.path.abspath('/System/Library/Fonts/Supplemental'), found)

    def test_safe_family_uses_available_vietnamese_macos_fallback(self):
        rows = [{
            'family': 'Helvetica Neue',
            'path': '/System/Library/Fonts/Helvetica.ttc',
            'index': 0,
            'viet': True,
        }]
        old_catalog, old_by_family = fonts._catalog, fonts._by_family
        try:
            fonts._catalog = rows
            fonts._by_family = {'helvetica neue': rows[0]}
            self.assertEqual(
                fonts.safe_family('Missing Windows Font', 'ĐỊA ĐIỂM'),
                'Helvetica Neue')
        finally:
            fonts._catalog, fonts._by_family = old_catalog, old_by_family


if __name__ == '__main__':
    unittest.main()
