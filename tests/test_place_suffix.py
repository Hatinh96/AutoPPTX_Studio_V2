"""Hậu tố vị trí GP trên tên ảnh (GPG/GPF/GPI…) không phải mã lệch."""
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autopptx2.datasource import (
    image_group_code, match_row, split_image_stem, strip_place_suffix,
    ImageLibrary,
)


class TestPlaceSuffix(unittest.TestCase):
    def test_screenshot_gpg(self):
        cases = (
            ('CDCONGNGHETHUDUCGPG', 'CDCONGNGHETHUDUC', 'GPG'),
            ('DHQUOCGIAHCMCS2GPG', 'DHQUOCGIAHCMCS2', 'GPG'),
            ('KTXDHCONGNGHEKYTHUATGPG', 'KTXDHCONGNGHEKYTHUAT', 'GPG'),
        )
        for raw, want, sfx in cases:
            base, got = strip_place_suffix(raw)
            self.assertEqual(base, want, raw)
            self.assertEqual(got.upper(), sfx)

    def test_gp_family_and_numbers(self):
        for raw, want in (
            ('SCHOOLGPF', 'SCHOOL'),
            ('SCHOOLGPI', 'SCHOOL'),
            ('SCHOOLGPP', 'SCHOOL'),
            ('SCHOOLGPS', 'SCHOOL'),
            ('SCHOOLGP', 'SCHOOL'),
            ('SCHOOLGPF1', 'SCHOOL'),
            ('SCHOOLGPG2', 'SCHOOL'),
            ('SCHOOLGP3', 'SCHOOL'),
            ('SCHOOL_GPF', 'SCHOOL'),
            ('SCHOOL.GPG', 'SCHOOL'),
            ('SCHOOL-GPI', 'SCHOOL'),
            ('SCHOOLGP.F', 'SCHOOL'),
            ('SCHOOLLCD', 'SCHOOL'),
        ):
            self.assertEqual(strip_place_suffix(raw)[0], want, raw)

    def test_keep_campus_cs2(self):
        self.assertEqual(strip_place_suffix('DHQUOCGIAHCMCS2')[0],
                         'DHQUOCGIAHCMCS2')

    def test_too_short_not_stripped(self):
        self.assertEqual(strip_place_suffix('ABCGP'), ('ABCGP', ''))

    def test_filename_and_copy_suffix(self):
        self.assertEqual(image_group_code('CDCONGNGHETHUDUCGPG.jpg'),
                         'CDCONGNGHETHUDUC')
        self.assertEqual(image_group_code('CDCONGNGHETHUDUCGPF (1).JPG'),
                         'CDCONGNGHETHUDUC')
        self.assertEqual(split_image_stem('CDCONGNGHETHUDUCGPF (2)'),
                         ('CDCONGNGHETHUDUC', 'GPF (2)'))

    def test_match_row_exact_not_fuzzy(self):
        by_code = {
            'CDCONGNGHETHUDUC': {'Code_RP': 'CDCONGNGHETHUDUC', 'Name': 'TĐ'},
        }
        row, key, kind = match_row('CDCONGNGHETHUDUCGPG', by_code)
        self.assertEqual(kind, 'exact')
        self.assertEqual(key, 'CDCONGNGHETHUDUC')
        self.assertEqual(row['Name'], 'TĐ')

    def test_groups_merge_gpg_and_gpf(self):
        with tempfile.TemporaryDirectory() as td:
            Path(td, 'CDCONGNGHETHUDUCGPG.jpg').write_bytes(b'x')
            Path(td, 'CDCONGNGHETHUDUCGPF.jpg').write_bytes(b'x')
            Path(td, 'CDCONGNGHETHUDUCGPF (1).jpg').write_bytes(b'x')
            lib = ImageLibrary()
            lib.set_folders([td], recursive=False)
            g = lib.groups()
            self.assertEqual(list(g), ['CDCONGNGHETHUDUC'])
            self.assertEqual(len(g['CDCONGNGHETHUDUC']), 3)


if __name__ == '__main__':
    unittest.main()
