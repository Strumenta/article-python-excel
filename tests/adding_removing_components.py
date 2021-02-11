import unittest
from openpyxl import Workbook

class TestAddRemoveSheets(unittest.TestCase):
    def testEmptySheetNames(self):
        wb = Workbook()
        self.assertEqual(['Sheet'], wb.sheetnames)
        wb.close()

    def testAddSheet(self):
        wb = Workbook()
        wb.create_sheet('Another Sheet', 0)
        self.assertEqual(['Another Sheet', 'Sheet'], wb.sheetnames)
        wb.close()

    def testRemoveSheet(self):
        wb = Workbook()
        del wb['Sheet']
        self.assertEqual([], wb.sheetnames)
        wb.close()

class TestAddRemoveCells(unittest.TestCase):
    def testAddRemoveRows(self):
        wb = Workbook()
        self.assertEqual(wb.active.max_row, 0)
        wb.active.insert_rows(0, 3)
        self.assertEqual(wb.active.max_row, 3)
        del wb.active['B']
        self.assertEqual(wb.active.max_row, 2)
        wb.close()
