import unittest
import os
from src.utils import parse_input_periods

def get_filename_without_extension(file_path):
    base_name = os.path.basename(file_path)
    name_without_extension = os.path.splitext(base_name)[0]
    return name_without_extension

def list_all_holiday_images(dir):
    directory_path = os.path.abspath(dir)
    if os.path.exists(directory_path):
        files = os.listdir(directory_path)
        file_names_without_extension = [get_filename_without_extension(file) 
                                        for file in files if file.endswith(".png")]
        holiday_images_without_extension = [int(file) for file in file_names_without_extension
                                            if file.isdigit()]
        holiday_images_without_extension.sort()
        return [str(i) for i in holiday_images_without_extension]
    else:
        print(f"Directory '{directory_path}' does not exist.")
        return []
    

class TestListAllHolidayImages(unittest.TestCase):

    def test_list_all_holiday_images(self):
        result = list_all_holiday_images('images')

        expected_result = [str(i) for i in range(0, 18)]
        self.assertEqual(result, expected_result)


if __name__ == '__main__':
    unittest.main()

