import unittest
from Dynamic_Cap_Library import *

class TestDynamicCapLibrary(unittest.TestCase):

    def test_placeDataIntoSheet(self):
        client = validate_credentials()

        sheet = client.open(DYNAMIC_CAP_FILE)

        sheet_instance = sheet.get_worksheet(0)

        sheet_instance.insert_rows([1, 2, 3, 4])



if __name__ == '__main__':
    unittest.main()