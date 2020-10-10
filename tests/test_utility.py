import unittest
import openpyxl
import os

from context import workbookOperations
from context import utility

TEST_SHEET = os.path.dirname(__file__) + '/data/test.xlsx'


class TestUtility(unittest.TestCase):
    def test_insertIntoDict(self):
        headerArray = ['first', 'second', 'third']

        result = utility.insertIntoDict({}, headerArray, "value")

        self.assertDictEqual(
            result, {'third': {'second': {'first': 'value'}}}, msg=result)

        result2 = utility.insertIntoDict(
            {'third': {}}, headerArray, "value")

        self.assertDictEqual(
            result2, {'third': {'second': {'first': 'value'}}}, msg=result)

        result3 = utility.insertIntoDict(
            {'third': {}, 'steve': 'foo'}, headerArray, "value")

        self.assertDictEqual(
            result3, {'third': {'second': {'first': 'value'}}, 'steve': 'foo'}, msg=result)


if __name__ == '__main__':
    unittest.main()
