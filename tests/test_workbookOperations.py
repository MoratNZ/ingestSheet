import unittest
import openpyxl

from context import workbookOperations
from config import TEST_SHEET


class TestWorkbookOperations(unittest.TestCase):
    def test_workbook_loads(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        self.assertIsInstance(book, openpyxl.workbook.workbook.Workbook)


if __name__ == '__main__':
    unittest.main()
