import unittest
from datetime import datetime, timedelta
from src.date_utils import last_day_of_month

def last_day_of_month(year, month):
    # Calculate the first day of the next month
    first_day_of_next_month = datetime(year, month, 1) + timedelta(days=32)

    # Subtract one day to get the last day of the current month
    last_day_of_month = first_day_of_next_month - timedelta(days=first_day_of_next_month.day)

    return last_day_of_month

class TestLastDayOfMonth(unittest.TestCase):

    def test_last_day_of_january(self):
        result = last_day_of_month(2023, 1)
        self.assertEqual(result, datetime(2023, 1, 31))

    def test_last_day_of_february_non_leap_year(self):
        result = last_day_of_month(2023, 2)
        self.assertEqual(result, datetime(2023, 2, 28))

    def test_last_day_of_february_leap_year(self):
        result = last_day_of_month(2024, 2)
        self.assertEqual(result, datetime(2024, 2, 29))

    def test_last_day_of_december(self):
        result = last_day_of_month(2023, 12)
        self.assertEqual(result, datetime(2023, 12, 31))

if __name__ == '__main__':
    unittest.main()
