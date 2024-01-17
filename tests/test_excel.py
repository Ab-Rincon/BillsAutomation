import unittest
from utils.excel import read_excel_data


class TestExcel(unittest.TestCase):

    def test_read_data(self):
        filekey = 'tests/input/test.xlsx'
        read_excel_data(filekey)


if __name__ == '__main__':
    unittest.main()
