
import unittest
import openpyxl
import os

from context import ingestSheet

TEST_SHEET = os.path.dirname(__file__) + '/data/test.xlsx'


class TestSomething(unittest.TestCase):
    def test_workbook_loads(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        self.assertIsInstance(book, openpyxl.workbook.workbook.Workbook)

    def test_isChildMergedCell(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        self.assertFalse(ingestSheet.isChildMergedCell(
            sheet.cell(column=1, row=1)))  # empty cell
        self.assertFalse(ingestSheet.isChildMergedCell(
            sheet.cell(column=2, row=1)))  # standard cell with content
        self.assertFalse(ingestSheet.isChildMergedCell(
            sheet.cell(column=3, row=1)))  # parent merged cell
        self.assertTrue(ingestSheet.isChildMergedCell(
            sheet.cell(column=4, row=1)))  # child merged cell

    def test_isEmptyCell(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        self.assertTrue(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=1)))  # empty cell
        self.assertTrue(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=2)))  # cell containing spaces
        self.assertFalse(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=3)))  # cell containing zero
        self.assertFalse(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=4)))  # cell containing non-zero number
        self.assertFalse(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=5)))  # cell containing text
        self.assertFalse(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=6)))  # cell containing formula
        self.assertFalse(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=7)))  # cell containing float
        self.assertFalse(ingestSheet.isEmptyCell(
            sheet.cell(column=13, row=8)))  # cell containing date

    def test_getCellValue(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        self.assertIsNone(ingestSheet.getCellValue(
            sheet.cell(row=1, column=1)))
        self.assertEqual(ingestSheet.getCellValue(
            sheet.cell(row=1, column=2)), "parent one")
        self.assertEqual(ingestSheet.getCellValue(
            sheet.cell(row=1, column=3)), "parent two")
        self.assertEqual(ingestSheet.getCellValue(
            sheet.cell(row=1, column=4)), "parent two")

    def test_getParseHeaders(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        headers = ingestSheet.parseHeaders(sheet, headerRowCount=2)

        self.assertEqual(len(headers), 9)

        # TODO need more tests

    def test_insertIntoDict(self):
        headerArray = ['first', 'second', 'third']

        result = ingestSheet.insertIntoDict({}, headerArray, "value")

        self.assertDictEqual(
            result, {'third': {'second': {'first': 'value'}}}, msg=result)

        result2 = ingestSheet.insertIntoDict(
            {'third': {}}, headerArray, "value")

        self.assertDictEqual(
            result2, {'third': {'second': {'first': 'value'}}}, msg=result)

        result3 = ingestSheet.insertIntoDict(
            {'third': {}, 'steve': 'foo'}, headerArray, "value")

        self.assertDictEqual(
            result3, {'third': {'second': {'first': 'value'}}, 'steve': 'foo'}, msg=result)

    def test_parseRow(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        headers = ingestSheet.parseHeaders(sheet, headerRowCount=2)

        row1 = sheet[3]
        result1 = ingestSheet.parseRow(headers, row1)
        self.assertDictEqual(result1, {'parent one': {'child one': 1}, 'parent two': {'child one': 3, 'child two': 6,
                                                                                      'child three': 9}, 'parent three': {'child one': 11, 'child two': 14}, 'no parent': 17, 'no children': 20})

        row2 = sheet[4]
        result2 = ingestSheet.parseRow(headers, row2)
        self.assertDictEqual(result2, {'parent one': {'child one': 2}, 'parent two': {'child one': 4, 'child two': 7,
                                                                                      'child three': None}, 'parent three': {'child one': 12, 'child two': 15}, 'no parent': 18, 'no children': 21})

    def test_parseSheet(self):
        book = ingestSheet.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        result = ingestSheet.parseSheet(sheet, headerRowCount=2)

        self.assertDictEqual(result, {'row one': {'parent one': {'child one': 1}, 'parent two': {'child one': 3, 'child two': 6, 'child three': 9}, 'parent three': {'child one': 11, 'child two': 14}, 'no parent': 17, 'no children': 20}, 'row two': {'parent one': {'child one': 2}, 'parent two': {'child one': 4, 'child two': 7, 'child three': None}, 'parent three': {'child one': 12, 'child two': 15}, 'no parent': 18, 'no children': 21}, 'row three': {
                             'parent one': {'child one': None}, 'parent two': {'child one': 5, 'child two': 7, 'child three': 10}, 'parent three': {'child one': 13, 'child two': 16}, 'no parent': 19, 'no children': 22}, None: {'parent one': {'child one': None}, 'parent two': {'child one': None, 'child two': None, 'child three': None}, 'parent three': {'child one': None, 'child two': None}, 'no parent': None, 'no children': None}})


if __name__ == '__main__':
    unittest.main()
