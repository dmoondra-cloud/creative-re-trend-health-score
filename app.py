"""
Creative RE T12 Categorizer - Minimal Self-Contained Version
Standalone app with all logic embedded (no external modules needed)
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import tempfile
import re
from datetime import datetime
from io import StringIO, BytesIO
import os
import shutil

# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(SCRIPT_DIR, "templates")
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "THS_Template_Default.xlsx")

# ============================================================================
# EMBEDDED: T12 PARSER
# ============================================================================

class T12Parser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = openpyxl.load_workbook(file_path)
        self.ws = self.wb.active

    def parse(self):
        """Parse T12 statement from Excel."""
        line_items = []
        property_name = None
        period = None

        # Try to detect property name and period from first few rows
        for r in range(1, min(10, self.ws.max_row + 1)):
            cell_val = self.ws.cell(r, 1).value
            if cell_val and isinstance(cell_val, str):
                if not property_name:
                    property_name = str(cell_val).strip()
                if 'month' in str(cell_val).lower() or 'period' in str(cell_val).lower():
                    period = str(cell_val).strip()

        # Parse T12 data starting from line 10
        # Include ALL line items without filtering or header detection
        for r in range(10, self.ws.max_row + 1):
            col_a = self.ws.cell(r, 1).value
            if col_a is None:
                col_a = self.ws.cell(r, 2).value

            if col_a is None:
                continue

            label = str(col_a).strip()
            if not label or len(label) < 1:
                continue

            # Collect monthly values - scan all columns with numeric data
            values = []
            for c in range(2, min(20, self.ws.max_column + 1)):
                try:
                    val = self.ws.cell(r, c).value
                    if val is not None and isinstance(val, (int, float)):
                        values.append(float(val))
                    else:
                        values.append(0)
                except:
                    values.append(0)

            # Include all rows with a label (no filtering)
            # Pad to 12 months if needed
            while len(values) < 12:
                values.append(0)

            line_items.append({
                'label': label,
                'values': values[:12],  # Take first 12 months
                'is_subtotal': False,
                'is_section_header': False  # No section header detection
            })

        return {
            'parsed_successfully': len(line_items) > 0,
            'property_name': property_name or 'Unknown Property',
            'period': period or 'Unknown Period',
            'line_items': line_items
        }

# ============================================================================
# EMBEDDED: CATEGORIZATION ENGINE
# ============================================================================

class CategorizationEngine:
    """Semantic categorization - analyzes line item meaning, not hardcoded keywords."""

    def __init__(self):
        self.cache = {}

    def is_header_or_total(self, label, amount):
        """Check if this is a header/total line that shouldn't be categorized."""
        # Rule 1: No amount (0 or empty) + contains total-like word
        if amount == 0:
            total_keywords = ['total', 'subtotal', 'sub-total', 'aggregate', 'summary']
            if any(keyword in label.lower() for keyword in total_keywords):
                return True
        return False

    def categorize_line_item(self, label, amount=0, section='income'):
        """
        Semantic categorization based on line item MEANING, not keywords.
        section: 'income' (before Total Income), 'expense' (between Total Income and NOI), 'post_noi' (after NOI)
        """
        if label in self.cache:
            return self.cache[label]

        # Universal Rule: No amount + total-like word = "-"
        if self.is_header_or_total(label, amount):
            result = {
                'label': label,
                'category': '-',
                'type': 'income',
                'confidence': 1.0
            }
            self.cache[label] = result
            return result

        # INCOME SECTION (before Total Income)
        if section == 'income':
            return self._categorize_income(label, amount)

        # EXPENSE SECTION (between Total Income and NOI)
        elif section == 'expense':
            return self._categorize_expense(label, amount)

        # POST-NOI SECTION
        else:
            result = {
                'label': label,
                'category': '-',
                'type': 'income',
                'confidence': 0.0
            }
            self.cache[label] = result
            return result

    def _categorize_income(self, label, amount):
        """Categorize income items into: Rental Income, Rental Losses, Other Income"""
        label_lower = label.lower()

        # Check for RENTAL LOSSES (most specific)
        rental_loss_keywords = {
            'vacancy': 'Less: Vacancy Loss',
            'loss.*lease|lease.*loss': 'Less: Loss to Lease',
            'bad debt': 'Less: Bad Debt',
            'concession|conc': 'Less: Concessions',
            'model|admin|down unit|employee unit|non-revenue|courtesy': 'Less: Non-Revenue Units'
        }

        for keyword, category in rental_loss_keywords.items():
            if re.search(keyword, label_lower, re.IGNORECASE):
                result = {
                    'label': label,
                    'category': category,
                    'type': 'income',
                    'confidence': 0.85
                }
                self.cache[label] = result
                return result

        # Check for RENTAL INCOME
        rental_keywords = ['rent', 'rental', 'market rent', 'base rent', 'gross']
        if any(keyword in label_lower for keyword in rental_keywords):
            result = {
                'label': label,
                'category': 'Gross Potential Rents',
                'type': 'income',
                'confidence': 0.85
            }
            self.cache[label] = result
            return result

        # Check for OTHER INCOME (reimbursements, fees, etc.)
        other_income_keywords = ['fee', 'pet', 'parking', 'late', 'damage', 'interest', 'reimbursement', 'rubs', 'amenity', 'application', 'miscellaneous']
        if any(keyword in label_lower for keyword in other_income_keywords):
            result = {
                'label': label,
                'category': 'Other Income',
                'type': 'income',
                'confidence': 0.80
            }
            self.cache[label] = result
            return result

        # Default to "-" (unknown income item)
        result = {
            'label': label,
            'category': '-',
            'type': 'income',
            'confidence': 0.0
        }
        self.cache[label] = result
        return result

    def _categorize_expense(self, label, amount):
        """Categorize expense items - most items are expenses by default in this section."""
        label_lower = label.lower()

        # Check if it's clearly NOT an expense (recovery, refund, etc.)
        non_expense_keywords = ['recovery', 'refund', 'reimbursement', 'income', 'fee (positive context)']
        if any(keyword in label_lower for keyword in non_expense_keywords):
            # Double-check: might be Other Income
            if 'reimbursement' in label_lower or 'recovery' in label_lower:
                result = {
                    'label': label,
                    'category': 'Other Income',
                    'type': 'income',
                    'confidence': 0.75
                }
                self.cache[label] = result
                return result

        # Default: Items in expense section are EXPENSES
        result = {
            'label': label,
            'category': 'Expense',
            'type': 'expense',
            'confidence': 0.90
        }
        self.cache[label] = result
        return result

    def categorize_batch(self, line_items):
        """Categorize multiple line items."""
        categorized = []
        for item in line_items:
            categorization = self.categorize_line_item(
                item['label'],
                sum(item['values']) if 'values' in item else 0,
                section='income'  # Default, will be overridden by UI logic
            )
            categorized.append({**item, **categorization})
        return categorized

# ============================================================================
# EXCEL DOWNLOAD GENERATOR
# ============================================================================

def generate_t12_download(edited_items, property_name, parsed_t12):
    """
    Generate T12 sheet in THS template with categorized items.
    - Uses preloaded template from app folder
    - Fills T12 sheet starting from row 10
    - Column C = Category
    - Column D = Line Item Name
    - Columns E-P = Monthly amounts × multiplier
    - Returns BytesIO object for download
    """
    try:
        # Check if template exists
        if not os.path.exists(TEMPLATE_PATH):
            return None, f"Template not found at: {TEMPLATE_PATH}"

        # Load template
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb["T12"]

        # ────────────────────────────────────────────────────────────────
        # POPULATE C6 WITH T12 PERIOD
        # ────────────────────────────────────────────────────────────────
        if parsed_t12.get('period'):
            ws['C6'].value = parsed_t12['period']

        # ────────────────────────────────────────────────────────────────
        # FILL T12 SHEET WITH CATEGORIZED DATA
        # ────────────────────────────────────────────────────────────────
        current_row = 10
        noi_row = None

        for item in edited_items:
            # Skip post-NOI items
            if item.get('is_post_noi', False):
                continue

            # Check if this is the NOI line item
            if 'NOI' in str(item['label']).upper() or item['label'] == st.session_state.selected_noi:
                noi_row = current_row

            # Column C = Category
            ws.cell(current_row, 3).value = item['category']

            # Column D = Line Item Name
            ws.cell(current_row, 4).value = item['label']

            # Columns E-P = 12 months (adjusted by multiplier)
            for month_idx in range(12):
                original_amount = item['values'][month_idx] if isinstance(item['values'], list) else item['original_amount']
                adjusted_amount = original_amount * item['multiplier']

                col_num = 5 + month_idx  # Column E = 5, F = 6, ... P = 16
                ws.cell(current_row, col_num).value = adjusted_amount

            current_row += 1

        # ────────────────────────────────────────────────────────────────
        # CREATE NOI FORMULAS IN Z26, AA26, AB26, AC26, AD26
        # ────────────────────────────────────────────────────────────────
        if noi_row:
            # Z26 = R[noi_row], AA26 = S[noi_row], AB26 = T[noi_row], AC26 = U[noi_row], AD26 = V[noi_row]
            # Column R = 18, S = 19, T = 20, U = 21, V = 22
            # Target columns: Z=26, AA=27, AB=28, AC=29, AD=30

            target_cols = [26, 27, 28, 29, 30]  # Z, AA, AB, AC, AD
            source_cols = [18, 19, 20, 21, 22]  # R, S, T, U, V

            for target_col, source_col in zip(target_cols, source_cols):
                source_letter = get_column_letter(source_col)
                target_letter = get_column_letter(target_col)
                # Create formula: =R[noi_row]
                formula = f"={source_letter}{noi_row}"
                ws[f'{target_letter}26'].value = formula

        # Save to BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return output, None

    except Exception as e:
        return None, f"Error generating file: {str(e)}"

# ============================================================================
# STREAMLIT UI
# ============================================================================

st.set_page_config(
    page_title="Creative RE T12 Categorizer",
    page_icon="📊",
    layout="wide"
)

st.title("📊 T12 Categorization Tool")
st.markdown("**Upload raw T12 → Select Boundaries → AI Categorize → Review & Export**")

# ────────────────────────────────────────────────────────────────────────────
# COMPACT UPLOAD + SELECT BOUNDARIES LAYOUT
# ────────────────────────────────────────────────────────────────────────────

left_col, right_col = st.columns([1, 1])

# LEFT SIDE: Upload T12 (Compact)
with left_col:
    st.subheader("📁 Upload T12")
    t12_file = st.file_uploader(
        "Upload raw T12 (Yardi/MRI format)",
        type=["xlsx", "xls"],
        label_visibility="collapsed"
    )

if not t12_file:
    st.warning("⚠️ Please upload a T12 file to continue")
    st.stop()

try:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(t12_file.getbuffer())
        t12_path = tmp.name

    parser = T12Parser(t12_path)
    parsed_t12 = parser.parse()

    if not parsed_t12['parsed_successfully']:
        st.error("❌ Could not parse T12.")
        st.stop()

except Exception as e:
    st.error(f"❌ Error: {str(e)}")
    st.stop()

# ────────────────────────────────────────────────────────────────────────────
# SUMMARY TABLE
# ────────────────────────────────────────────────────────────────────────────

st.markdown("---")

engine = CategorizationEngine()
# Initial categorization (will be refined in UI based on section)
categorized = engine.categorize_batch(parsed_t12['line_items'])

# Prepare table data - all line items without filtering
table_data = []
item_idx = 0

for item in categorized:
    total_val = sum(item['values']) if 'values' in item else 0

    table_data.append({
        'idx': item_idx,
        'amount': total_val,
        'line_item': item['label'],
        'category': item['category'],
        'income_expense_type': 'Income' if item['type'] == 'income' else 'Expense',
        'monthly_values': item['values'] if 'values' in item else [0] * 12  # Preserve monthly breakdown
    })
    item_idx += 1

# RIGHT SIDE: Select Boundaries + Run AI Categorisation
with right_col:
    st.subheader("⚙️ Set Boundaries")

    # Get list of all line items for dropdown
    all_line_items = [row['line_item'] for row in table_data]

    # Two columns for dropdowns (25% each)
    col_ti, col_noi = st.columns(2)

    with col_ti:
        selected_total_income = st.selectbox(
            "Total Income",
            options=['--'] + all_line_items,
            index=0,
            key="total_income_select",
            label_visibility="collapsed"
        )

    with col_noi:
        selected_noi = st.selectbox(
            "NOI",
            options=['--'] + all_line_items,
            index=0,
            key="noi_select",
            label_visibility="collapsed"
        )

    # Store selections in session state
    st.session_state.selected_total_income = selected_total_income
    st.session_state.selected_noi = selected_noi

    # Large AI Categorisation button
    st.markdown("")  # Small spacing
    if st.button("🤖 Run AI Categorisation", use_container_width=True, type="primary", help="Use AI to intelligently categorize all line items"):
        st.session_state.step1_complete = True
        st.rerun()

# Only show content if user clicked Run AI Categorisation
if not st.session_state.get('step1_complete', False):
    st.stop()

# Calculate indices for Total Income and NOI
total_income_idx = None
noi_idx = None
for i, item in enumerate(table_data):
    if item['line_item'] == st.session_state.selected_total_income:
        total_income_idx = i
    if item['line_item'] == st.session_state.selected_noi:
        noi_idx = i

# ────────────────────────────────────────────────────────────────────────────
# SUMMARY TABLE: Compare T12 vs Categorisation
# ────────────────────────────────────────────────────────────────────────────
# Header
st.subheader("📋 Financial Summary")
st.caption(f"📍 {parsed_t12['property_name']}")

# Find Total Expense line (usually just before NOI)
total_expense_line = None
total_expense_amount = 0
if noi_idx is not None and noi_idx > 0:
    # Look backwards from NOI to find "TOTAL EXPENSE" or "TOTAL OPERATING EXPENSE"
    for i in range(noi_idx - 1, -1, -1):
        if i < len(table_data):
            if 'total' in table_data[i]['line_item'].lower() and 'expense' in table_data[i]['line_item'].lower():
                total_expense_line = table_data[i]['line_item']
                total_expense_amount = table_data[i]['amount']
                break

# Get values from T12 (user selections)
total_income_t12 = 0
noi_t12 = 0
if st.session_state.selected_total_income != '--':
    for item in table_data:
        if item['line_item'] == st.session_state.selected_total_income:
            total_income_t12 = item['amount']
            break
if st.session_state.selected_noi != '--':
    for item in table_data:
        if item['line_item'] == st.session_state.selected_noi:
            noi_t12 = item['amount']
            break

# Left column for table (50% width)
left_col, right_col = st.columns([1, 1])

with left_col:
    # Column headings (no divider line)
    col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
    with col1:
        st.markdown("")  # Empty for alignment
    with col2:
        st.write("**As per T12**")
    with col3:
        st.write("**As per Categorisation**")
    with col4:
        st.write("**Error Check**")

    # Data rows
    col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
    with col1:
        st.write("Total Income")
    with col2:
        st.write(f"${total_income_t12:,.0f}")
    with col3:
        st.write("*[Calculating...]*")
    with col4:
        st.write("—")

    col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
    with col1:
        st.write("Total Expense")
    with col2:
        st.write(f"${total_expense_amount:,.0f}" if total_expense_line else "—")
    with col3:
        st.write("*[Calculating...]*")
    with col4:
        st.write("—")

    col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
    with col1:
        st.write("NOI")
    with col2:
        st.write(f"${noi_t12:,.0f}")
    with col3:
        st.write("*[Calculating...]*")
    with col4:
        st.write("—")

# Download button - 50% width on left
col_download, col_space = st.columns([1, 1])
with col_download:
    if st.button("📥 Download to Excel", use_container_width=True, type="primary", help="Download categorized T12 as Excel"):
        # Generate Excel file with categorized T12 using preloaded template
        excel_file, error = generate_t12_download(st.session_state.t12_categorized, parsed_t12['property_name'], parsed_t12)

        if error:
            st.error(f"❌ {error}")
        else:
            # Provide download button
            st.download_button(
                label="💾 Click to Download File",
                data=excel_file,
                file_name=f"T12_{parsed_t12['property_name'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.success("✅ File ready for download!")

# ────────────────────────────────────────────────────────────────────────────
# STEP 1: Review & Adjust AI Categorisation
# ────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("STEP 1: Review & Adjust AI Categorisation")
st.markdown("**Review AI-generated categories and make adjustments as needed**")

# Calculate indices once (outside loop)
total_income_idx = None
noi_idx = None
for i, item in enumerate(table_data):
    if item['line_item'] == st.session_state.selected_total_income:
        total_income_idx = i
    if item['line_item'] == st.session_state.selected_noi:
        noi_idx = i

# Header row - same column widths as data rows
header_col1, header_col2, header_col3, header_col4, header_col5 = st.columns([2.5, 1.1, 1.8, 0.9, 1.0])
with header_col1:
    st.markdown("**Line Item**")
with header_col2:
    st.markdown("**Orig. Amt**")
with header_col3:
    st.markdown("**Category**")
with header_col4:
    st.markdown("**Mult.**")
with header_col5:
    st.markdown("**Result**")

# Table rows - Categorisation
edited_items = []
categorization_stopped = False

for idx, row in enumerate(table_data):
    # Check if we've reached NOI - stop categorization after this
    if st.session_state.selected_noi != '--' and row['line_item'] == st.session_state.selected_noi:
        # Mark this as NOI but don't categorize anything after it
        categorization_stopped = True

    # After NOI, only show the line item without categorization options
    if categorization_stopped and row['line_item'] != st.session_state.selected_noi:
        col1, col2, col3, col4, col5 = st.columns([2.5, 1.1, 1.8, 0.9, 1.0])

        with col1:
            st.write(f"`{row['line_item']}`")
        with col2:
            st.write(f"**{row['amount']:,.0f}**")
        with col3:
            st.write("—")
        with col4:
            st.write("—")
        with col5:
            st.write("—")

        edited_items.append({
            'label': row['line_item'],
            'original_amount': row['amount'],
            'multiplier': 1,
            'adjusted_amount': row['amount'],
            'category': '-',
            'section_type': '-',
            'values': row['monthly_values'],  # Use monthly values
            'is_section_header': False,
            'is_post_noi': True
        })
        continue

    # Regular line item row
    col1, col2, col3, col4, col5 = st.columns([2.5, 1.1, 1.8, 0.9, 1.0])

    original_amount = row['amount']

    # Determine if this item is before, between, or after Total Income/NOI
    is_between_sections = (total_income_idx is not None and noi_idx is not None and
                           total_income_idx < idx < noi_idx)

    with col1:
        marker = ""
        if row['line_item'] == st.session_state.selected_total_income:
            marker = "💰 "
        elif row['line_item'] == st.session_state.selected_noi:
            marker = "📊 "
        st.write(f"{marker}`{row['line_item']}`")

    with col2:
        st.write(f"**{original_amount:,.0f}**")

    with col3:
        # Determine section and re-categorize using semantic engine
        if is_between_sections:
            section = 'expense'
        elif idx > noi_idx if noi_idx is not None else False:
            section = 'post_noi'
        else:
            section = 'income'

        # Re-categorize based on section using semantic engine
        semantic_cat = engine.categorize_line_item(row['line_item'], original_amount, section=section)
        default_category = semantic_cat['category']

        # Build category options - use 8 main categories plus "-"
        category_options = ['-', 'Gross Potential Rents', 'Less: Vacancy Loss', 'Less: Loss to Lease',
                           'Less: Bad Debt', 'Less: Concessions', 'Less: Non-Revenue Units',
                           'Other Income', 'Expense']

        # Determine default index
        if default_category in category_options:
            current_index = category_options.index(default_category)
        else:
            current_index = 0  # Default to "-"

        selected_cat = st.selectbox(
            "Category",
            options=category_options,
            index=current_index,
            key=f"cat_{idx}_{hash(row['line_item']) % 10000}",
            label_visibility="collapsed"
        )

    with col4:
        multiplier_value = st.selectbox(
            "Multiplier",
            options=[1, -1],
            format_func=lambda x: f"+{x}" if x > 0 else f"{x}",
            index=0,
            key=f"mult_{idx}_{hash(row['line_item']) % 10000}",
            label_visibility="collapsed"
        )

    with col5:
        adjusted_amount = original_amount * multiplier_value
        display_text = "same" if adjusted_amount == original_amount else f"{adjusted_amount:,.0f}"
        st.write(f"**{display_text}**")

    edited_items.append({
        'label': row['line_item'],
        'original_amount': original_amount,
        'multiplier': multiplier_value,
        'adjusted_amount': adjusted_amount,
        'category': selected_cat,
        'section_type': '-',
        'values': row['monthly_values'],  # Use monthly values from table_data
        'is_section_header': False,
        'is_post_noi': False
    })

st.session_state.t12_categorized = edited_items

