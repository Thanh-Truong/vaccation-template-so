import unittest
from datetime import date, timedelta
from src.date_utils import generate_monthly_ranges_within_year
 

class TestGenerateMonthlyRangesWithinYear(unittest.TestCase):

    def test_ranges_within_single_year(self):
        start_date = date(2023, 3, 15)
        end_date = date(2023, 8, 25)

        result = generate_monthly_ranges_within_year(start_date, end_date)

        # Assuming that find_interval_intersection is correct
        expected_result = [
            (date(2023, 3, 1), date(2023, 3, 31)),
            (date(2023, 4, 1), date(2023, 4, 30)),
            (date(2023, 5, 1), date(2023, 5, 31)),
            (date(2023, 6, 1), date(2023, 6, 30)),
            (date(2023, 7, 1), date(2023, 7, 31)),
            (date(2023, 8, 1), date(2023, 8, 25))
        ]

        self.assertEqual(result, expected_result)

    def test_ranges_across_multiple_years(self):
        start_date = date(2022, 11, 15)
        end_date = date(2023, 2, 10)

        # Ensure that the function raises a ValueError
        with self.assertRaises(ValueError):
            generate_monthly_ranges_within_year(start_date, end_date)

if __name__ == '__main__':
    unittest.main()
