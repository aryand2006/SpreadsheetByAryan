import unittest
from src.engine.formulas import sum_formula, average_formula

class TestFormulas(unittest.TestCase):

    def test_sum_formula(self):
        self.assertEqual(sum_formula([1, 2, 3]), 6)
        self.assertEqual(sum_formula([-1, 1, 0]), 0)
        self.assertEqual(sum_formula([]), 0)

    def test_average_formula(self):
        self.assertEqual(average_formula([1, 2, 3]), 2)
        self.assertEqual(average_formula([1, 2, 3, 4]), 2.5)
        self.assertEqual(average_formula([]), 0)

if __name__ == '__main__':
    unittest.main()