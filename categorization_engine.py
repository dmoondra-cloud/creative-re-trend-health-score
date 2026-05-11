"""
Flexible categorization engine for T12 line items.
Handles pattern-based matching and special cases like RUBS,
concessions, rental losses, and reimbursements.
"""

import re
from typing import Dict, List, Tuple, Optional


class CategorizationEngine:
    """Categorizes T12 line items based on patterns and rules."""

    # Define category patterns - order matters (more specific first)
    CATEGORY_RULES = {
        'GPR (Gross Potential Rent)': {
            'patterns': [r'\bGross\s+Potential\s+Rent\b', r'\bGPR\b', r'^\s*Gross Rental Income$'],
            'type': 'income',
            'section': 'Gross Rental Income'
        },
        'Concessions': {
            'patterns': [r'Rent\s+Concessions?\b', r'Concession', r'Lease Concession'],
            'type': 'income',
            'section': 'Rental Loss'
        },
        'Vacancy Loss': {
            'patterns': [r'Vacancy\s+Loss', r'Vacant Unit', r'Unoccupied'],
            'type': 'income',
            'section': 'Rental Loss'
        },
        'Delinquency Loss': {
            'patterns': [r'Delinquency\s+Loss', r'Bad Debt.*Rent', r'Collection Loss'],
            'type': 'income',
            'section': 'Rental Loss'
        },
        'Rental Credits/Deductions': {
            'patterns': [r'Rental\s+Credits', r'Rental\s+Deductions', r'Credit.*Rent'],
            'type': 'income',
            'section': 'Rental Loss'
        },
        'Gain/Loss to Old Lease': {
            'patterns': [r'Gain.*Old\s+Lease', r'Loss.*Old\s+Lease'],
            'type': 'income',
            'section': 'Rental Income Adjustments'
        },
        'Prepaid Rent': {
            'patterns': [r'Prepaid\s+Rent', r'Deferred\s+Rent'],
            'type': 'income',
            'section': 'Rental Income Adjustments'
        },
        'Employee Rent Discounts': {
            'patterns': [r'Employee\s+Rent\s+Discount', r'Staff\s+Discount'],
            'type': 'income',
            'section': 'Rental Loss'
        },
        'Military/Special Discounts': {
            'patterns': [r'Military\s+.*Discount', r'Special\s+.*Discount', r'Employer\s+Discount'],
            'type': 'income',
            'section': 'Rental Loss'
        },
        'Damages': {
            'patterns': [r'Damage\s+Charge', r'Damages\s+Recovered', r'Damage Revenue'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Late Fees': {
            'patterns': [r'Late\s+Fee', r'Late Payment', r'Delinquent Fee'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Application Fees': {
            'patterns': [r'Application\s+Fee', r'Application Income'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Pet Fees': {
            'patterns': [r'Pet\s+Fee', r'Pet\s+Rent', r'Animal.*Fee'],
            'type': 'income',
            'section': 'Other Income'
        },
        'NSF Fees': {
            'patterns': [r'NSF\s+Fee', r'Insufficient\s+Fund', r'Returned Check'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Lease/Termination Fees': {
            'patterns': [r'Termination\s+Fee', r'Lease.*Fee', r'Lease Violation'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Other Fee Income': {
            'patterns': [r'Other\s+Fee', r'Miscellaneous\s+Fee', r'Admin\s+Fee'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Utility Income': {
            'patterns': [r'\bUtility\s+Income\b', r'Utility\s+Recharge', r'Utility\s+Billing'],
            'type': 'income',
            'section': 'Other Income'
        },
        'RUBS (Resident Utility Billing)': {
            'patterns': [r'\bRUBS\b', r'Resident.*Utility', r'Utility.*Resident'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Reimbursed Expenses': {
            'patterns': [r'Reimburse', r'Reimbursement', r'Recovery'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Insurance Reimbursement': {
            'patterns': [r'Insurance.*Reimburse', r'Reimburse.*Insurance'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Tax Reimbursement': {
            'patterns': [r'Tax.*Reimburse', r'Reimburse.*Tax'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Parking': {
            'patterns': [r'Parking', r'Parking Revenue'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Laundry': {
            'patterns': [r'Laundry', r'Washer.*Dryer'],
            'type': 'income',
            'section': 'Other Income'
        },
        'Vending': {
            'patterns': [r'Vending', r'Vending Machine'],
            'type': 'income',
            'section': 'Other Income'
        },
        # EXPENSE CATEGORIES
        'Property Management': {
            'patterns': [r'Property.*Management', r'Management\s+Fee'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Payroll': {
            'patterns': [r'Wages', r'Salary', r'Payroll', r'Labor', r'Staff'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Utilities': {
            'patterns': [r'Electricity', r'Water.*Sewer', r'Gas.*Heat', r'Utility Expense'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Contract Services': {
            'patterns': [r'Landscaping', r'Pest Control', r'Trash', r'Contract'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Repairs & Maintenance': {
            'patterns': [r'Repair', r'Maintenance', r'HVAC', r'Plumbing', r'Electrical',
                        r'Painting', r'Appliance'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Advertising': {
            'patterns': [r'Advertising', r'Marketing', r'Leasing'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Taxes': {
            'patterns': [r'Property\s+Tax', r'Income\s+Tax', r'Payroll\s+Tax'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Insurance': {
            'patterns': [r'\bInsurance\b', r'Liability'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
        'Bad Debt': {
            'patterns': [r'Bad\s+Debt', r'Allowance.*Doubtful'],
            'type': 'expense',
            'section': 'Operating Expenses'
        },
    }

    def __init__(self):
        self.categorization_cache = {}

    def categorize_line_item(self, line_label: str, line_value: float = 0) -> Dict:
        """
        Categorize a single line item based on pattern matching.
        Returns category info and any special handling flags.
        """

        # Check cache first
        if line_label in self.categorization_cache:
            return self.categorization_cache[line_label]

        result = {
            'line_label': line_label,
            'category': 'Other Income',  # Default
            'type': 'income',
            'section': 'Other Income',
            'is_special_case': False,
            'special_handling': None,
            'confidence': 0.0
        }

        # Special case: negative values in expense section = income
        if line_value < 0 and any(
            keyword in line_label.lower()
            for keyword in ['utility', 'rubs', 'reimburs', 'recovery']
        ):
            result['is_special_case'] = True
            result['special_handling'] = 'negate_value'
            result['category'] = 'Other Income'
            result['type'] = 'income'

        # Pattern matching
        best_match = None
        best_score = 0

        for category, rules in self.CATEGORY_RULES.items():
            for pattern in rules['patterns']:
                match = re.search(pattern, line_label, re.IGNORECASE)
                if match:
                    # Score based on match position and length
                    score = 100 - (match.start() / len(line_label) * 20)
                    if score > best_score:
                        best_score = score
                        best_match = {
                            'category': category,
                            'type': rules['type'],
                            'section': rules['section'],
                            'confidence': min(1.0, score / 100)
                        }

        if best_match:
            result.update(best_match)

        # Cache the result
        self.categorization_cache[line_label] = result
        return result

    def categorize_batch(self, line_items: List[Dict]) -> List[Dict]:
        """Categorize multiple line items."""
        categorized = []
        for item in line_items:
            categorization = self.categorize_line_item(
                item['label'],
                sum(item['values']) if 'values' in item else 0
            )
            categorized.append({**item, **categorization})
        return categorized

    def validate_categorization(self, categorized_items: List[Dict]) -> Dict:
        """Validate categorization results and flag items for review."""
        issues = {
            'low_confidence': [],
            'suspicious_items': [],
            'unmatched_items': [],
            'validation_warnings': []
        }

        for item in categorized_items:
            if item.get('confidence', 0) < 0.6:
                issues['low_confidence'].append({
                    'label': item['label'],
                    'suggested_category': item['category'],
                    'confidence': item['confidence']
                })

            # Flag suspicious patterns
            if 'income' in item['label'].lower() and item['type'] == 'expense':
                issues['suspicious_items'].append({
                    'label': item['label'],
                    'issue': 'Possible miscategorization: Income listed under expenses'
                })

        return issues

    def get_category_mapping(self) -> Dict[str, List[str]]:
        """Return a mapping of categories to their patterns."""
        return {
            category: rules['patterns']
            for category, rules in self.CATEGORY_RULES.items()
        }

    def add_custom_rule(self, category: str, patterns: List[str], rule_type: str, section: str):
        """Add custom categorization rules."""
        if category not in self.CATEGORY_RULES:
            self.CATEGORY_RULES[category] = {
                'patterns': patterns,
                'type': rule_type,
                'section': section
            }
        else:
            self.CATEGORY_RULES[category]['patterns'].extend(patterns)
        # Clear cache when rules change
        self.categorization_cache.clear()
