"""
Rent Roll processor - handles RR loading, column detection, and analysis.
"""

import pandas as pd
import openpyxl
from typing import Dict, List, Optional


class RRProcessor:
    """Processes Rent Roll data with intelligent column mapping."""

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.rr_data = None
        self.units_summary = None
        self.column_mappings = {}

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
                        self._detect_columns()
                        return df
                except:
                    continue

            if df is None:
                # Fallback: read first sheet
                df = pd.read_excel(self.file_path)
                self.rr_data = df
                self._detect_columns()
                return df

        except Exception as e:
            raise ValueError(f"Could not load Rent Roll: {str(e)}")

    def _detect_columns(self) -> Dict:
        """Auto-detect common column mappings."""
        if self.rr_data is None:
            return {}

        columns = list(self.rr_data.columns)

        # Define patterns for column detection
        patterns = {
            'unit_number': ['unit', 'unit no', 'unit #', 'unit number'],
            'unit_type': ['type', 'floorplan', 'floor plan', 'bd', 'br', 'bedroom'],
            'sqft': ['sqft', 'sq ft', 'square feet', 'size'],
            'resident_name': ['tenant', 'resident', 'name', 'lessee'],
            'market_rent': ['market', 'market rent', 'market rate'],
            'actual_rent': ['rent', 'actual rent', 'lease rent', 'monthly rent'],
            'status': ['status', 'occupancy', 'occupied'],
            'lease_start': ['lease start', 'start date', 'move in'],
            'lease_end': ['lease end', 'end date', 'move out']
        }

        # Auto-detect columns
        detected = {}
        for field, keywords in patterns.items():
            for col in columns:
                col_lower = col.lower()
                for keyword in keywords:
                    if keyword in col_lower:
                        detected[field] = col
                        break
                if field in detected:
                    break

        self.column_mappings = detected
        return detected

    def get_column_suggestions(self) -> Dict:
        """Get auto-detected and all available columns for mapping."""
        if self.rr_data is None:
            self.load_rr()

        available_columns = list(self.rr_data.columns)

        return {
            'detected': self.column_mappings,
            'available': available_columns,
            'header_row': 0
        }

    def get_summary(self) -> Dict:
        """Get summary statistics from Rent Roll."""
        if self.rr_data is None:
            self.load_rr()

        df = self.rr_data

        summary = {
            'total_units': len(df),
            'columns': list(df.columns),
            'data': df,
            'column_mappings': self.column_mappings
        }

        # Try to extract occupancy stats
        status_col = self.column_mappings.get('status')
        if status_col and status_col in df.columns:
            summary['occupancy_stats'] = df[status_col].value_counts().to_dict()
        else:
            # Try fallback
            for col in ['Status', 'status', 'Unit Status', 'Occupancy']:
                if col in df.columns:
                    summary['occupancy_stats'] = df[col].value_counts().to_dict()
                    break

        return summary

    def get_gpr_from_rr(self) -> float:
        """Extract GPR from Rent Roll data using mapped columns."""
        if self.rr_data is None:
            self.load_rr()

        # Look for mapped rent columns
        rent_col = self.column_mappings.get('actual_rent')
        if rent_col and rent_col in self.rr_data.columns:
            try:
                return pd.to_numeric(self.rr_data[rent_col], errors='coerce').sum()
            except:
                pass

        # Fallback: search all columns
        rent_columns = [col for col in self.rr_data.columns
                       if 'rent' in col.lower() or 'rate' in col.lower()]

        if rent_columns:
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
        expected_cols = ['unit_number', 'unit_type', 'status', 'actual_rent']
        for col_field in expected_cols:
            if col_field not in self.column_mappings:
                issues['missing_columns'].append(col_field)

        # Check for empty rows
        empty_count = self.rr_data.isna().sum().sum()
        if empty_count > len(self.rr_data) * 0.1:  # More than 10% empty
            issues['warnings'].append(
                f"Data sparsity: {empty_count} empty cells detected"
            )

        return issues
