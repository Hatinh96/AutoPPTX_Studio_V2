import os
import tempfile
import unittest
from collections import OrderedDict

from openpyxl import Workbook

from autopptx2.datasource import (
    ExcelSource,
    inspect_excel,
    looks_like_order_list,
    order_groups,
    sort_codes_by_list_order,
)


class TestListOrder(unittest.TestCase):
    def test_filename_and_sheet_detect_list(self):
        self.assertTrue(looks_like_order_list(r'C:\x\List CF.xlsx'))
        self.assertTrue(looks_like_order_list('University List.xlsx'))
        self.assertTrue(looks_like_order_list('master.xlsx', ['Coffee Shop & Milk Tea List']))
        self.assertFalse(looks_like_order_list('Standard_Master_All_Channels.xlsx'))
        self.assertFalse(looks_like_order_list('master.xlsx', ['Data', 'Summary']))

    def test_sort_follows_excel_row_order(self):
        rows = [
            {'Code_RP': 'HLCALMETTE'},
            {'Code_RP': 'HLLIBERTY'},
            {'Code_RP': 'HL119HAMNGHI'},
        ]
        by_code = {r['Code_RP']: r for r in rows}
        self.assertEqual(
            sort_codes_by_list_order(
                ['HL119HAMNGHI', 'HLCALMETTE', 'HLLIBERTY'], rows, by_code),
            ['HLCALMETTE', 'HLLIBERTY', 'HL119HAMNGHI'])

    def test_order_groups_list_mode(self):
        groups = OrderedDict([
            ('HL119HAMNGHI', ['c.jpg']),
            ('HLCALMETTE', ['a.jpg']),
        ])
        rows = [{'Code_RP': 'HLCALMETTE'}, {'Code_RP': 'HL119HAMNGHI'}]
        by_code = {r['Code_RP']: r for r in rows}
        out = order_groups(groups, by_code, excel_rows=rows, sort_mode='list')
        self.assertEqual(list(out), ['HLCALMETTE', 'HL119HAMNGHI'])

    def test_uploaded_list_overrides_master_order(self):
        groups = OrderedDict([
            ('HLLIBERTY', ['b.jpg']),
            ('HLCALMETTE', ['a.jpg']),
            ('HL119HAMNGHI', ['c.jpg']),
        ])
        master = [
            {'Code_RP': 'HL119HAMNGHI'},
            {'Code_RP': 'HLCALMETTE'},
            {'Code_RP': 'HLLIBERTY'},
        ]
        uploaded = [
            {'Code_RP': 'HLCALMETTE'},
            {'Code_RP': 'HLLIBERTY'},
            {'Code_RP': 'HL119HAMNGHI'},
        ]
        by_code = {r['Code_RP']: r for r in master}
        out = order_groups(
            groups, by_code, excel_rows=master, sort_mode='list',
            order_rows=uploaded)
        self.assertEqual(list(out), ['HLCALMETTE', 'HLLIBERTY', 'HL119HAMNGHI'])

    def _write_list_book(self):
        wb = Workbook()
        summary = wb.active
        summary.title = 'Summary'
        summary.append(['Code_RP', 'Name'])
        summary.append(['SHOULD_SKIP', 'Summary row'])
        coffee = wb.create_sheet('Coffee Shop & Milk Tea List')
        coffee.append(['No.', 'Report Code', 'Name'])
        coffee.append(['1', 'HLCALMETTE', 'Highlands Calmette'])
        coffee.append(['2', 'HLLIBERTY', 'Highlands Liberty'])
        fd, path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        wb.save(path)
        wb.close()
        return path

    def test_load_uses_list_sheet_not_summary(self):
        path = self._write_list_book()
        try:
            src = ExcelSource()
            self.assertTrue(src.load(path), src.error)
            self.assertEqual(
                [r['Code_RP'] for r in src.rows],
                ['HLCALMETTE', 'HLLIBERTY'])
            self.assertTrue(looks_like_order_list(path, src.source_sheets))
            info = inspect_excel(path)
            self.assertTrue(info.get('prefer_list_order'))
            self.assertEqual(info.get('n'), 2)
        finally:
            if os.path.exists(path):
                os.remove(path)


if __name__ == '__main__':
    unittest.main()
