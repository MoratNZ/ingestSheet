
import unittest
import openpyxl
import os

from context import workbookOperations
from context import sheetOperations

TEST_SHEET = os.path.dirname(__file__) + '/data/test.xlsx'


class TestSheetOperations(unittest.TestCase):
    def test_parseHeaders(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        headers = sheetOperations.parseHeaders(sheet, headerRowCount=2)

        self.assertEqual(len(headers), 9)

        # TODO need more tests

    def test_parseSheet(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        result = sheetOperations.parseSheet(sheet, headerRowCount=2)

        self.assertDictEqual(result, {'row one': {'parent one': {'child one': 1}, 'parent two': {'child one': 3, 'child two': 6, 'child three': 9}, 'parent three': {'child one': 11, 'child two': 14}, 'no parent': 17, 'no children': 20}, 'row two': {'parent one': {'child one': 2}, 'parent two': {'child one': 4, 'child two': 7, 'child three': None}, 'parent three': {'child one': 12, 'child two': 15}, 'no parent': 18, 'no children': 21}, 'row three': {
                             'parent one': {'child one': None}, 'parent two': {'child one': 5, 'child two': 7, 'child three': 10}, 'parent three': {'child one': 13, 'child two': 16}, 'no parent': 19, 'no children': 22}, None: {'parent one': {'child one': None}, 'parent two': {'child one': None, 'child two': None, 'child three': None}, 'parent three': {'child one': None, 'child two': None}, 'no parent': None, 'no children': None}})


if __name__ == '__main__':
    unittest.main()
