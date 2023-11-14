import unittest
from src.utils import parse_input_periods

class TestParseInputPeriods(unittest.TestCase):
    def test_valid_input(self):
        input_string = "1 - 5,  7-10, 15-20"
        result = parse_input_periods(input_string)
        expected = [(1, 5), (7, 10), (15, 20)]
        self.assertEqual(result, expected)

    def test_invalid_input(self):
        input_string = "1 - abc, 7-10, 15-20"
        with self.assertRaises(ValueError):
            parse_input_periods(input_string)

if __name__ == '__main__':
    unittest.main()
