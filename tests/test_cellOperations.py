import unittest
from context import cellOperations
from context import workbookOperations
from config import TEST_SHEET


class TestCellOperations(unittest.TestCase):
    def test_isChildMergedCell(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        self.assertFalse(cellOperations.isChildMergedCell(
            sheet.cell(column=1, row=1)))  # empty cell
        self.assertFalse(cellOperations.isChildMergedCell(
            sheet.cell(column=2, row=1)))  # standard cell with content
        self.assertFalse(cellOperations.isChildMergedCell(
            sheet.cell(column=3, row=1)))  # parent merged cell
        self.assertTrue(cellOperations.isChildMergedCell(
            sheet.cell(column=4, row=1)))  # child merged cell

    def test_isEmptyCell(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        self.assertTrue(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=1)))  # empty cell
        self.assertTrue(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=2)))  # cell containing spaces
        self.assertFalse(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=3)))  # cell containing zero
        self.assertFalse(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=4)))  # cell containing non-zero number
        self.assertFalse(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=5)))  # cell containing text
        self.assertFalse(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=6)))  # cell containing formula
        self.assertFalse(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=7)))  # cell containing float
        self.assertFalse(cellOperations.isEmptyCell(
            sheet.cell(column=13, row=8)))  # cell containing date

    def test_getCellValue(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet = book['Sheet1']

        self.assertIsNone(cellOperations.getCellValue(
            sheet.cell(row=1, column=1)))
        self.assertEqual(cellOperations.getCellValue(
            sheet.cell(row=1, column=2)), "parent one")
        self.assertEqual(cellOperations.getCellValue(
            sheet.cell(row=1, column=3)), "parent two")
        self.assertEqual(cellOperations.getCellValue(
            sheet.cell(row=1, column=4)), "parent two")


if __name__ == '__main__':
    unittest.main()
