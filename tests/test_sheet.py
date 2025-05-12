import unittest
from pyspreadsheet.core.sheet import Sheet
from pyspreadsheet.core.cell import Cell

class TestSheet(unittest.TestCase):

    def setUp(self):
        self.sheet = Sheet("Test Sheet")

    def test_add_cell(self):
        cell = Cell(value=10)
        self.sheet.add_cell(cell, row=0, col=0)
        self.assertEqual(self.sheet.get_cell(0, 0).value, 10)

    def test_remove_cell(self):
        cell = Cell(value=20)
        self.sheet.add_cell(cell, row=1, col=1)
        self.sheet.remove_cell(row=1, col=1)
        self.assertIsNone(self.sheet.get_cell(1, 1))

    def test_get_cell_data(self):
        cell = Cell(value=30)
        self.sheet.add_cell(cell, row=2, col=2)
        self.assertEqual(self.sheet.get_cell_data(row=2, col=2), 30)

    def test_add_multiple_cells(self):
        for i in range(3):
            for j in range(3):
                self.sheet.add_cell(Cell(value=i + j), row=i, col=j)
        
        self.assertEqual(self.sheet.get_cell(0, 0).value, 0)
        self.assertEqual(self.sheet.get_cell(1, 1).value, 2)
        self.assertEqual(self.sheet.get_cell(2, 2).value, 4)

if __name__ == '__main__':
    unittest.main()