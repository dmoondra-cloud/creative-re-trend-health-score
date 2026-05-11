"""
Rent Roll processor - handles RR loading and analysis.
"""

import pandas as pd
import openpyxl
from typing import Dict, List, Optional


class RRProcessor:
    """Processes Rent Roll data."""

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.rr_data = None
        self.units_summary = None

    def load_rr(self) -> pd.DataFrame:
        """Load Rent Roll from Excel file."""
        try:
            # Try common sheet names
            sheet_names = ['Rent Roll', 'RR', 'Units', 'Sheet1']
            df = None

            for sheet in sheet_names:
                try:
                    df = pd.read_excel(self.file_path, sheet_name=sheet)
                    if len(df) > 0:
                        self.rr_data = df
                        return df
                except:
                    continue

            if df is None:
                # Fallback: read first sheet
                df = pd.read_excel(self.file_path)
                self.rr_data = df
                return df

        except Exception as e:
            raise ValueError(f"Could not load Rent Roll: {str(e)}")

    def get_summary(self) -> Dict:
        """Get summary statistics from Rent Roll."""
        if self.rr_data is None:
            self.load_rr()

        df = self.rr_data

        summary = {
            'total_units': len(df),
            'columns': list(df.columns),
            'data': df
        }

        # Try to extract unit status
        status_column = None
        for col in ['Status', 'status', 'Unit Status', 'Occupancy']:
            if col in df.columns:
                status_column = col
                break

        if status_column:
            summary['occupancy_stats'] = df[status_column].value_counts().to_dict()

        return summary

    def get_gpr_from_rr(self) -> float:
        """Extract GPR from Rent Roll data."""
        if self.rr_data is None:
            self.load_rr()

        # Look for rent columns
        rent_columns = [col for col in self.rr_data.columns
                       if 'rent' in col.lower() or 'rate' in col.lower()]

        if rent_columns:
            # Sum all rent columns
            total_gpr = 0
            for col in rent_columns:
                try:
                    total_gpr += pd.to_numeric(
                        self.rr_data[col], errors='coerce'
                    ).sum()
                except:
                    pass
            return total_gpr

        return 0

    def validate_rr(self) -> Dict:
        """Validate Rent Roll data quality."""
        issues = {
            'missing_columns': [],
            'empty_rows': [],
            'data_type_issues': [],
            'warnings': []
        }

        if self.rr_data is None:
            self.load_rr()

        # Check for expected columns
        expected_cols = ['Unit No', 'Unit Type', 'Status', 'Rent']
        for col in expected_cols:
            if not any(col.lower() in str(c).lower() for c in self.rr_data.columns):
                issues['missing_columns'].append(col)

        # Check for empty rows
        empty_count = self.rr_data.isna().sum().sum()
        if empty_count > len(self.rr_data) * 0.1:  # More than 10% empty
            issues['warnings'].append(
                f"Data sparsity: {empty_count} empty cells detected"
            )

        return issues
