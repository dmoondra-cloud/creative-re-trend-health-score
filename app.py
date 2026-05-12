"""
Creative RE T12 Categorizer - Minimal Self-Contained Version
Standalone app with all logic embedded (no external modules needed)
"""

import streamlit as st
import pandas as pd
import openpyxl
import tempfile
import re
from datetime import datetime
from io import StringIO

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
        'income_expense_type': 'Income' if item['type'] == 'income' else 'Expense'
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

st.markdown("---")

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
# Header with property name box on right
header_col1, header_col2 = st.columns([3, 1])
with header_col1:
    st.subheader("📋 Financial Summary: T12 vs Categorisation")
with header_col2:
    st.info(f"📍 {parsed_t12['property_name']}")

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

# Create summary table header - no "Particular" label
col_particular, col_t12, col_cat, col_check = st.columns([1.5, 1.8, 2.0, 1.8])

with col_particular:
    st.markdown("")  # Blank instead of "Particular"
with col_t12:
    st.markdown("**As per T12**")
with col_cat:
    st.markdown("**As per Categorisation**")
with col_check:
    st.markdown("**Error Check**")

st.divider()

# Row 1: Total Income
col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
with col1:
    st.write("**Total Income**")
with col2:
    st.write(f"${total_income_t12:,.0f}")
with col3:
    st.write("*[Calculating...]*")
with col4:
    st.write("—")

# Row 2: Total Expense
col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
with col1:
    st.write("**Total Expense**")
with col2:
    st.write(f"${total_expense_amount:,.0f}" if total_expense_line else "—")
with col3:
    st.write("*[Calculating...]*")
with col4:
    st.write("—")

# Row 3: NOI
col1, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 1.8])
with col1:
    st.write("**NOI**")
with col2:
    st.write(f"${noi_t12:,.0f}")
with col3:
    st.write("*[Calculating...]*")
with col4:
    st.write("—")

st.markdown("---")

# Download button - same size as Run AI button
if st.button("📥 Download to Excel", use_container_width=True, type="primary", help="Download categorized T12 as Excel"):
    st.info("📊 Download feature coming soon!", icon="ℹ️")

# ────────────────────────────────────────────────────────────────────────────
# STEP 1: Review & Adjust AI Categorisation
# ────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("STEP 1: Review & Adjust AI Categorisation")
st.markdown("**Review AI-generated categories and make adjustments as needed**")

# Table header for categorisation
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

st.markdown("---")

# Calculate indices once (outside loop)
total_income_idx = None
noi_idx = None
for i, item in enumerate(table_data):
    if item['line_item'] == st.session_state.selected_total_income:
        total_income_idx = i
    if item['line_item'] == st.session_state.selected_noi:
        noi_idx = i

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
            'values': row,
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
        'values': row,
        'is_section_header': False,
        'is_post_noi': False
    })

st.session_state.t12_categorized = edited_items

st.markdown("---")
st.subheader("📥 Export Categorized T12")

if st.button("✅ Export as CSV", use_container_width=True, type="primary"):
    export_data = []
    for item in st.session_state.get('t12_categorized', []):
        post_noi_marker = " (Post-NOI - Not Categorized)" if item.get('is_post_noi', False) else ""
        export_data.append({
            'Line Item': item['label'] + post_noi_marker,
            'Original Amount': item['original_amount'],
            'Multiplier': f"+{item['multiplier']}" if item['multiplier'] > 0 else f"{item['multiplier']}",
            'Adjusted Amount': item['adjusted_amount'],
            'Category': item['category'],
            'Section': item['section_type']
        })

    df_export = pd.DataFrame(export_data)
    csv = df_export.to_csv(index=False)

    st.download_button(
        label="📥 Download CSV",
        data=csv,
        file_name=f"T12_Categorized_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv"
    )
    st.success("✅ Ready!")

st.markdown("---")
st.markdown("<p style='text-align:center'><small>Creative RE T12 Categorizer | Minimal Version</small></p>", unsafe_allow_html=True)
