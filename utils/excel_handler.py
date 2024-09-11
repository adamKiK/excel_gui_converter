"""
This module declares the FileHandler abstract class and the ExcelHandler class.
The FileHandler class is an abstract class that defines the interface for all
file handlers. The ExcelHandler class is a concrete class that implements the
FileHandler interface for Excel files.
"""

import os
from typing import List, Optional
from abc import ABC, abstractmethod
import pandas as pd
from openpyxl import Workbook


class FileHandler(ABC):
    @abstractmethod
    def read(self, input_path: str) -> pd.DataFrame:
        pass

    @abstractmethod
    def write(self, data: pd.DataFrame, output_path: str) -> None:
        pass


class ExcelHandler(FileHandler):
    def __init__(self, file_path: str):
        self.file_path = self.set_file_path(file_path)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def read(self, sheet_name: Optional[str] = None) -> pd.DataFrame:
        try:
            return pd.read_excel(self.file_path, sheet_name=sheet_name)
        except Exception as e:
            raise IOError(f"Error reading Excel file: {e}")

    def write(self, data: pd.DataFrame, sheet_name: str = 'Sheet1', if_sheet_exists: str = 'replace') -> None:
        try:
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a' if self.file_exists else 'w') as writer:
                data.to_excel(writer, sheet_name=sheet_name, index=False, if_sheet_exists=if_sheet_exists)
        except Exception as e:
            raise IOError(f"Error writing to Excel file: {e}")

    @staticmethod
    def set_file_path(file_path: str) -> str:
        if not file_path:
            raise ValueError('File path is empty')
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            raise ValueError('File path must be an Excel file (.xlsx or .xls)')
        return file_path

    def get_sheet_names(self) -> List[str]:
        try:
            return pd.ExcelFile(self.file_path).sheet_names
        except Exception as e:
            raise IOError(f"Error getting sheet names: {e}")

    def create_file_if_not_exists(self) -> None:
        if not self.file_exists:
            Workbook().save(self.file_path)
            print(f"Created new Excel file: {self.file_path}")

    @property
    def file_exists(self) -> bool:
        return os.path.exists(self.file_path)
