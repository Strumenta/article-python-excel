import unittest
import formulas
from openpyxl import Workbook, load_workbook
from openpyxl.cell import Cell


class TestFormulas(unittest.TestCase):
    def test_data_only(self):
        wb = load_workbook('data/test_formulas.xlsx', data_only=True)
        self.assertEqual(wb.active['A1'].value, 42)
        self.assertEqual(wb.active['B1'].value, 21)
        self.assertEqual(wb.active['C1'].value, 31.5)
        wb.close()

    def test_calculate(self):
        xl_model = formulas.ExcelModel().loads('data/test_formulas.xlsx').finish()
        solution = xl_model.calculate()
        c1 = solution['\'[TEST_FORMULAS.XLSX]SHEET1\'!C1']
        self.assertEqual(c1.value[0], 31.5)

    def test_compute_formula(self):
        wb = load_workbook('data/test_formulas.xlsx')
        self.assertEqual(compute_formula(wb.active['C1']), 31.5)
        wb.active['A1'].value = 100
        self.assertEqual(compute_formula(wb.active['C1']), 75)
        wb.close()


def compute_formula(cell: Cell):
    value = cell.value
    sheet = cell.parent
    if not isinstance(value, str) or not value.startswith('='):
        return value
    func = formulas.Parser().ast(value)[1].compile()
    args = []
    for cell in func.inputs.keys():
        args.append(compute_formula(sheet[cell]))
    return func(*args)
