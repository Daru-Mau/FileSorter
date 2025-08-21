#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel data model for representing Excel spreadsheet data
"""

import pandas as pd
from typing import Dict, List, Any, Optional, Tuple


class ExcelModel:
    """Model for representing Excel data"""

    def __init__(self, file_path: str = None):  # type: ignore
        """
        Initialize an Excel model

        Args:
            file_path: Path to the Excel file (optional)
        """
        self.file_path = file_path
        self.data = None
        self.columns = []

        # Column mappings
        self.name_column = ""
        self.id_column = ""
        self.date_column = ""

        if file_path:
            self.load_data(file_path)

    def load_data(self, file_path: str) -> bool:
        """
        Load data from Excel file

        Args:
            file_path: Path to the Excel file

        Returns:
            True if loading was successful, False otherwise
        """
        try:
            self.data = pd.read_excel(file_path)
            self.columns = list(self.data.columns)
            self.file_path = file_path

            # Try to guess column mappings
            self._guess_column_mappings()

            return True
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")
            return False

    def _guess_column_mappings(self) -> None:
        """Guess column mappings based on column names"""
        if not self.columns:
            return

        # Try to find name, ID, and date columns
        name_matches = [col for col in self.columns if 'name' in col.lower()]
        id_matches = [col for col in self.columns if 'id' in col.lower()]
        date_matches = [col for col in self.columns if 'date' in col.lower()]

        if name_matches:
            self.name_column = name_matches[0]
        elif self.columns:
            self.name_column = self.columns[0]

        if id_matches:
            self.id_column = id_matches[0]
        elif len(self.columns) > 1:
            self.id_column = self.columns[1]

        if date_matches:
            self.date_column = date_matches[0]

    def get_row_as_dict(self, row_idx: int) -> Dict[str, Any]:
        """
        Get a specific row as a dictionary

        Args:
            row_idx: Row index

        Returns:
            Dictionary with column names as keys and row values as values
        """
        if self.data is None or row_idx >= len(self.data):
            return {}

        row = self.data.iloc[row_idx]
        return {col: row[col] for col in self.columns}

    def find_match(self, filename: str) -> Tuple[bool, Optional[int], Optional[Dict[str, Any]]]:
        """
        Find a matching row for a filename

        Args:
            filename: Filename to match

        Returns:
            Tuple containing:
            - Boolean indicating if a match was found
            - Row index of the match (or None)
            - Row data as dictionary (or None)
        """
        if self.data is None or not self.name_column:
            return False, None, None

        # Remove extension from filename for matching
        filename_without_ext = filename
        if '.' in filename:
            filename_without_ext = filename.rsplit('.', 1)[0]

        # Try exact match first
        for idx, row in self.data.iterrows():
            excel_value = str(row[self.name_column])

            if excel_value == filename_without_ext or excel_value == filename:
                return True, idx, self.get_row_as_dict(idx)  # type: ignore

        # Try partial match if exact match failed
        for idx, row in self.data.iterrows():
            excel_value = str(row[self.name_column])

            if filename_without_ext in excel_value or excel_value in filename_without_ext:
                return True, idx, self.get_row_as_dict(idx)  # type: ignore

        return False, None, None
