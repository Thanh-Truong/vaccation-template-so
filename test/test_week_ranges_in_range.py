import unittest
from datetime import datetime, timedelta
from src.date_utils import week_ranges_in_range

class TestWeekRangesInRange(unittest.TestCase):

    def test_week_ranges_within_month(self):
        start_date = datetime(2023, 11, 1)
        end_date = datetime(2023, 11, 30)
        result = week_ranges_in_range(start_date, end_date)

        # Assuming that find_interval_intersection is correct
        expected_result = [(44, (datetime(2023, 11, 1), datetime(2023, 11, 5))),
                           (45, (datetime(2023, 11, 6), datetime(2023, 11, 12))),
                           (46, (datetime(2023, 11, 13), datetime(2023, 11, 19))),
                           (47, (datetime(2023, 11, 20), datetime(2023, 11, 26))),
                           (48, (datetime(2023, 11, 27), datetime(2023, 11, 30)))]

        self.assertEqual(result, expected_result)

    def test_week_ranges_across_month(self):
        start_date = datetime(2023, 11, 28)
        end_date = datetime(2023, 12, 5)
        result = week_ranges_in_range(start_date, end_date)
        print(result)
        # Assuming that find_interval_intersection is correct
        expected_result = [(48, (datetime(2023, 11, 28), datetime(2023, 12, 3))),
                           (49, (datetime(2023, 12, 4), datetime(2023, 12, 5)))]

        self.assertEqual(result, expected_result)

    def test_week_ranges_across_year(self):
        start_date = datetime(2023, 12, 28)
        end_date = datetime(2024, 1, 10)

        # Ensure that the function raises a ValueError
        with self.assertRaises(ValueError):
            week_ranges_in_range(start_date, end_date)

if __name__ == '__main__':
    unittest.main()
