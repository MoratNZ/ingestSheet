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

        with self.assertRaises(ValueError):
            result4 = utility.insertIntoDict({}, [], "bob")

    def test_makeCamelCase(self):
        self.assertEqual(utility.makeCamelCase("Hello world"), "helloWorld")
        self.assertEqual(utility.makeCamelCase("Hello"), "hello")
        self.assertEqual(utility.makeCamelCase("helloWorld"), "helloworld")
        self.assertEqual(utility.makeCamelCase("hello_world"), "helloWorld")
        self.assertEqual(utility.makeCamelCase("hello"), "hello")
        self.assertEqual(utility.makeCamelCase("Hello"), "hello")
        self.assertEqual(utility.makeCamelCase("HeLlO"), "hello")


if __name__ == '__main__':
    unittest.main()
