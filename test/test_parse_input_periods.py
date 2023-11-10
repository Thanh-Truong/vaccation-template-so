import unittest
from src.utils import parse_input_periods

class TestParseInputPeriods(unittest.TestCase):

    def test_valid_periods(self):
        custom_periods = "1 - 4,  5 - 8, 9 - 12"
        result = parse_input_periods(custom_periods)

        expected_result = [(1, 4), (5, 8), (9, 12)]
        self.assertEqual(result, expected_result)

    def test_valid_periods_with_spaces(self):
        custom_periods = "  1  -  4 , 5  -  8 ,  9  -  12  "
        result = parse_input_periods(custom_periods)

        expected_result = [(1, 4), (5, 8), (9, 12)]
        self.assertEqual(result, expected_result)

    def test_valid_periods_mixed_format(self):
        custom_periods = "1 -4, 5-  8 ,  9-  12"
        result = parse_input_periods(custom_periods)

        expected_result = [(1, 4), (5, 8), (9, 12)]
        self.assertEqual(result, expected_result)

    def test_valid_periods_with_text(self):
        custom_periods = "1-4, 5-8, 9-12, invalid"
        result = parse_input_periods(custom_periods)
        expected_result = [(1, 4), (5, 8), (9, 12)]
        self.assertEqual(result, expected_result)

    def test_empty_input(self):
        custom_periods = ""
        result = parse_input_periods(custom_periods)
        self.assertEqual(result, [])

if __name__ == '__main__':
    unittest.main()
