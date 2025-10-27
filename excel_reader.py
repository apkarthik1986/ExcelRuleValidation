"""
Excel Reader Module
Handles reading and parsing Excel files for rule validation.
"""
import openpyxl
import pandas as pd
from typing import Dict, List, Any, Optional


class ExcelReader:
    """
    Excel file reader that provides access to data for validation.
    """
    
    def __init__(self, file_path: str):
        """
        Initialize the Excel reader.
        
        Args:
            file_path: Path to the Excel file
        """
        self.file_path = file_path
        self.workbook = None
        self.data_frame = None
        
    def load(self, sheet_name: Optional[str] = None, header_row: Optional[int] = 0) -> pd.DataFrame:
        """
        Load Excel file into a pandas DataFrame.
        
        Args:
            sheet_name: Name of the sheet to load (default: first sheet)
            
        Returns:
            DataFrame containing the Excel data
        """
        try:
            # pandas read_excel header parameter expects row index (0-based).
            # header_row is 0-based here; pass through to pandas so users can
            # choose which row contains column names (e.g. header_row=1 -> second row).
            if sheet_name:
                self.data_frame = pd.read_excel(self.file_path, sheet_name=sheet_name, header=header_row)
            else:
                self.data_frame = pd.read_excel(self.file_path, header=header_row)
            return self.data_frame
        except Exception as e:
            raise Exception(f"Error loading Excel file: {str(e)}")
    
    def get_columns(self) -> List[str]:
        """
        Get list of column names from the loaded data.
        
        Returns:
            List of column names
        """
        if self.data_frame is None:
            raise ValueError("No data loaded. Call load() first.")
        return self.data_frame.columns.tolist()
    
    def get_data(self) -> pd.DataFrame:
        """
        Get the loaded DataFrame.
        
        Returns:
            DataFrame containing the data
        """
        if self.data_frame is None:
            raise ValueError("No data loaded. Call load() first.")
        return self.data_frame
    
    def get_row_count(self) -> int:
        """
        Get the number of rows in the loaded data.
        
        Returns:
            Number of rows
        """
        if self.data_frame is None:
            raise ValueError("No data loaded. Call load() first.")
        return len(self.data_frame)
    
    def get_sheet_names(self) -> List[str]:
        """
        Get list of all sheet names in the Excel file.
        
        Returns:
            List of sheet names
        """
        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except Exception as e:
            raise Exception(f"Error reading sheet names: {str(e)}")
