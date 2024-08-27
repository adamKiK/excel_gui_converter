"""
This module is used to read and write spreadsheet workbooks, target specific
worksheets, and convert the data to a specific format.
"""

import os
import pandas as pd


#TODO: Add a method to check the path validity
#TODO: Add a method to get the available sheet names
#TODO:


class WorkbookHandler:
    """
    A class used to represent an Excel workbook.
    """
    # TODO: Add a sheetnames attribute
    def __init__(self, path: str):
        if not path.endswith('.xlsx') and not path.endswith('.xls'):
            raise ValueError('Only Excel workbooks are supported')
        elif not os.path.exists(path):
            raise FileNotFoundError(f'File not found: {path}')
        else:
            self.path = path


    def get_sheet_names(self):
        """
        Returns the names of the worksheets in the workbook.
        """


    def read(self, sheet: str) -> pd.DataFrame:
        """
        Reads a worksheet from the Excel workbook and returns it as a DataFrame.

        :param sheet: the name of the worksheet
        :return: the worksheet as a DataFrame
        """
        if not os.path.exists(self.path):
            raise FileNotFoundError(f'File not found: {self.path}')
        return pd.read_excel(self.path, sheet)

    def write(self, df, sheet: str):
        """
        Writes a DataFrame to the Excel workbook.

        :param df: the DataFrame to write
        :param sheet: the name of the worksheet
        """
        if not os.path.exists(self.path):
            df.to_excel(self.path, sheet_name=sheet, index=False)
        else:
            book = load_workbook(self.path)
            writer = pd.ExcelWriter(self.path, engine='openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            df.to_excel(writer, sheet_name=sheet, index=False)
            writer.save()
