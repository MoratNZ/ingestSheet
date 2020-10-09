
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


if __name__ == '__main__':
    unittest.main()
