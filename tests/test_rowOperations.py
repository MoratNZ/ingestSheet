
import unittest
import openpyxl
import os

from context import workbookOperations
from context import rowOperations
from context import sheetOperations

TEST_SHEET = os.path.dirname(__file__) + '/data/test.xlsx'


class TestRowOperations(unittest.TestCase):
    def test_parseRow(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        headers = sheetOperations.parseHeaders(sheet, headerRowCount=2)

        row1 = sheet[3]
        result1 = rowOperations.parseRow(headers, row1)
        self.assertDictEqual(result1, {'parent one': {'child one': 1}, 'parent two': {'child one': 3, 'child two': 6,
                                                                                      'child three': 9}, 'parent three': {'child one': 11, 'child two': 14}, 'no parent': 17, 'no children': 20})

        row2 = sheet[4]
        result2 = rowOperations.parseRow(headers, row2)
        self.assertDictEqual(result2, {'parent one': {'child one': 2}, 'parent two': {'child one': 4, 'child two': 7,
                                                                                      'child three': None}, 'parent three': {'child one': 12, 'child two': 15}, 'no parent': 18, 'no children': 21})


if __name__ == '__main__':
    unittest.main()
