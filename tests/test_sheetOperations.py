
import unittest
import openpyxl
import os

from context import workbookOperations
from context import sheetOperations

TEST_SHEET = os.path.dirname(__file__) + '/data/test.xlsx'


class TestSheetOperations(unittest.TestCase):
    def test_parseHeaders(self):
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet1 = book['Sheet1']

        headers = sheetOperations.parseHeaders(sheet1, lastHeaderRow=2)

        self.assertEqual(len(headers), 9, msg=headers)
        self.assertCountEqual(headers, [['row number'], ['child one', 'parent one'], ['child one', 'parent two'], ['child two', 'parent two'], [
                              'child three', 'parent two'], ['child one', 'parent three'], ['child two', 'parent three'], ['no parent'], ['no children']])

        headers1a = sheetOperations.parseHeaders(
            sheet1, lastHeaderRow=2, maxColumnGap=0)

        self.assertEqual(len(headers1a), 9, msg=headers)
        self.assertCountEqual(headers1a, [['row number'], ['child one', 'parent one'], ['child one', 'parent two'], ['child two', 'parent two'], [
                              'child three', 'parent two'], ['child one', 'parent three'], ['child two', 'parent three'], ['no parent'], ['no children']])

        sheet2 = book['Sheet2']
        headers2 = sheetOperations.parseHeaders(
            sheet2, firstHeaderRow=4, lastHeaderRow=5)

        self.assertEqual(len(headers2), 9, msg=headers2)
        self.assertCountEqual(headers2, [['row number'], ['child one', 'parent one'], ['child one', 'parent two'], ['child two', 'parent two'], [
                              'child three', 'parent two'], ['child one', 'parent three'], ['child two', 'parent three'], ['no parent'], ['no children']])
        headers2a = sheetOperations.parseHeaders(
            sheet2, firstHeaderRow=4, lastHeaderRow=5, maxColumnGap=0)

        self.assertEqual(len(headers2a), 9, msg=headers2)
        self.assertCountEqual(headers2a, [['row number'], ['child one', 'parent one'], ['child one', 'parent two'], ['child two', 'parent two'], [
                              'child three', 'parent two'], ['child one', 'parent three'], ['child two', 'parent three'], ['no parent'], ['no children']])
        headers3 = sheetOperations.parseHeaders(
            sheet2, firstHeaderRow=4, lastHeaderRow=5, camelCaseHeaders=True)

        self.assertEqual(len(headers3), 9, msg=headers3)
        self.assertCountEqual(headers3, [['rowNumber'], ['childOne', 'parentOne'], ['childOne', 'parentTwo'], ['childTwo', 'parentTwo'], [
                              'childThree', 'parentTwo'], ['childOne', 'parentThree'], ['childTwo', 'parentThree'], ['noParent'], ['noChildren']])

        sheet3 = book['Sheet3']
        headers4 = sheetOperations.parseHeaders(
            sheet3, firstHeaderRow=4, lastHeaderRow=5)
        self.assertEqual(len(headers4), 9, msg=headers4)
        self.assertCountEqual(headers4, [['row number', 'parent one'], ['child one', 'parent one'], ['child one', 'parent two'], ['child two', 'parent two'], [
                              'child three', 'parent two'], ['child one', 'parent three'], ['child two', 'parent three'], ['no parent'], ['no children']])

    def test_parseSheet(self):
        expectedResult = {
            'row one': {'row number': 'row one', 'parent one': {'child one': 1}, 'parent two': {'child one': 3, 'child two': 6, 'child three': 9}, 'parent three': {'child one': 11, 'child two': 14}, 'no parent': 17, 'no children': 20},
            'row two': {'row number': 'row two', 'parent one': {'child one': 2}, 'parent two': {'child one': 4, 'child two': 7, 'child three': None}, 'parent three': {'child one': 12, 'child two': 15}, 'no parent': 18, 'no children': 21},
            'row three': {'row number': 'row three', 'parent one': {'child one': None}, 'parent two': {'child one': 5, 'child two': 7, 'child three': 10}, 'parent three': {'child one': 13, 'child two': 16}, 'no parent': 19, 'no children': 22}
        }

        self.maxDiff = None
        book = workbookOperations.openWorkbook(TEST_SHEET)

        sheet1 = book['Sheet1']

        result = sheetOperations.parseSheet(sheet1, lastHeaderRow=2)

        self.assertDictEqual(expectedResult, result,
                             msg="\n\nresult is:\n{}\n\n".format(result))
        sheet2 = book['Sheet2']

        result2 = sheetOperations.parseSheet(
            sheet2, firstHeaderRow=4, lastHeaderRow=5)

        self.assertDictEqual(expectedResult, result2,
                             msg="\n\nresult is:\n{}\n\n".format(result2))


if __name__ == '__main__':
    unittest.main()
