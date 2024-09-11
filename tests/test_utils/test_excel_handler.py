"""
This test case checks the ExcelHandler class from utils/excel_handler.py.
"""
import os
import glob
import pathlib
import unittest
from utils.excel_handler import ExcelHandler

CWD = os.path.dirname(os.getcwd())
TEST_DATA_PATH = pathlib.PurePath(CWD).joinpath('tests', 'data')
EXCEL_INPUT_FILES = glob.glob(f'{TEST_DATA_PATH.joinpath("*.xls")}')


class TestExcelHandler(unittest.TestCase):
    def test_read_valid(self):
        pass


if __name__ == '__main__':
    unittest.main()
