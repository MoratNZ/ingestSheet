# ingestSheet

A python module to ingest the contents of an appropriately formatted Excel sheet as a dict.

The motivation for this is a stepping stone between config lists held in Excel, and ansible configuration.

## Expected sheet format

Assuming that you have a sheet example.xlsx formatted like so:

![expeected sheet format](https://github.com/MoratNZ/ingestSheet/blob/master/docs/spreadsheet_example.png)

There are two header rows (more or less can be used as required), with higher rows being parents of lower rows, and cell merges being used to show which children are associateed with which parents.

Below the header rows are three data rows, where the first column contains the label for that row.

The most basic use of ingestSheet involves opening an Excel workbook, selecting a worksheet, and then importing the contents of it:

```
>>> import ingestSheet
>>> book = ingestSheet.openWorkbook("example.xlsx")
>>> sheet1 = book['Sheet1']
>>> ingestSheet.parseSheet(sheet1, headerRowCount=2)
```

This returns:

```
 'row one': {'no children': 20,
             'no parent': 17,
             'parent one': {'child one': 1},
             'parent three': {'child one': 11, 'child two': 14},
             'parent two': {'child one': 3, 'child three': 9, 'child two': 6}},
 'row three': {'no children': 22,
               'no parent': 19,
               'parent one': {'child one': None},
               'parent three': {'child one': 13, 'child two': 16},
               'parent two': {'child one': 5,
                              'child three': 10,
                              'child two': 7}},
 'row two': {'no children': 21,
             'no parent': 18,
             'parent one': {'child one': 2},
             'parent three': {'child one': 12, 'child two': 15},
             'parent two': {'child one': 4,
                            'child three': None,
                            'child two': 7}}}
```
