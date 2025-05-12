import unittest
from pyspreadsheet.src.core.cell import Cell

class TestCell(unittest.TestCase):

    def setUp(self):
        self.cell = Cell()

    def test_initial_value(self):
        self.assertEqual(self.cell.value, None)

    def test_set_value(self):
        self.cell.value = 10
        self.assertEqual(self.cell.value, 10)

    def test_set_formula(self):
        self.cell.formula = "=SUM(1, 2)"
        self.assertEqual(self.cell.formula, "=SUM(1, 2)")

    def test_get_value_with_formula(self):
        self.cell.formula = "=SUM(1, 2)"
        self.cell.value = 3  # Assuming the formula evaluates to 3
        self.assertEqual(self.cell.evaluate(), 3)

if __name__ == '__main__':
    unittest.main()