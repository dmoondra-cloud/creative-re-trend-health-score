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
    """Categorizes T12 line items into 8 THS categories."""

    CATEGORY_RULES = {
        'Gross Potential Rents': {
            'patterns': [r'\bGross\s+(?:Market|Potential)?\s+Rent', r'\bGPR\b', r'Gross\s+Rental\s+Income', r'Market\s+Rent'],
            'type': 'income',
            'priority': 100
        },
        'Less: Vacancy Loss': {
            'patterns': [r'\bVacancy\b', r'Vacant\s+Unit', r'Unoccupied', r'Vacancy\s+Loss'],
            'type': 'income',
            'priority': 90
        },
        'Less: Loss to Lease': {
            'patterns': [r'\bLoss\b.*\bLease\b', r'Contract\s+Gain.*Loss', r'\bLTL\b'],
            'type': 'income',
            'priority': 90
        },
        'Less: Bad Debt': {
            'patterns': [r'\bBad\s+Debt\b(?!\s+Recovery)', r'Allowance.*Doubtful', r'Uncollectible'],
            'type': 'income',
            'priority': 85
        },
        'Less: Concessions': {
            'patterns': [r'Concession', r'Conc\s*-', r'Move[\s\-]?in\s+Special', r'Rent\s+Reduction', r'Rent\s+Concession'],
            'type': 'income',
            'priority': 85
        },
        'Less: Non-Revenue Units': {
            'patterns': [r'Model\s+Unit', r'Admin\s+Unit', r'Office\s+Unit', r'Down\s+Unit', r'Employee\s+Unit', r'Non[\s\-]?Revenue', r'Courtesy.*Discount'],
            'type': 'income',
            'priority': 85
        },
        'Other Income': {
            'patterns': [r'\bOther\s+(?:Property\s+)?Income', r'Pet\s+(?:Fee|Rent)', r'Parking', r'Laundry', r'Late\s+(?:Fee|Charge)', r'Application\s+Fee', r'Amenity\s+Fee', r'Reimbursement', r'RUBS', r'Damage\s+(?:Fee|Income)', r'Interest\s+Income', r'(?:Key|Access)\s+(?:Fee|Card)', r'Lease\s+Violation', r'Legal\s+(?:Fee|Collection)', r'NSF', r'Storage', r'Miscellaneous\s+Income', r'Administrative\s+Fee'],
            'type': 'income',
            'priority': 70
        }
    }

    def __init__(self):
        self.cache = {}

    def categorize_line_item(self, label, value=0):
        """Categorize a single line item based on semantic matching."""
        if label in self.cache:
            return self.cache[label]

        # Check if this is a header/total line that should be "-"
        has_empty_amount = (value == 0 and 'total' in label.lower())
        if has_empty_amount:
            result = {
                'label': label,
                'category': '-',
                'type': 'income',
                'confidence': 1.0
            }
            self.cache[label] = result
            return result

        # Try to match against category rules
        best_match = None
        best_score = 0
        best_priority = -1

        for category, rules in self.CATEGORY_RULES.items():
            priority = rules.get('priority', 50)

            for pattern in rules['patterns']:
                match = re.search(pattern, label, re.IGNORECASE)
                if match:
                    # Score based on match position and priority
                    position_score = 100 - (match.start() / max(len(label), 1) * 30)
                    total_score = position_score + priority

                    if total_score > best_score or (total_score == best_score and priority > best_priority):
                        best_score = total_score
                        best_priority = priority
                        best_match = {
                            'category': category,
                            'type': rules['type'],
                            'confidence': min(1.0, best_score / 150)
                        }

        # Default to "-" if no match found (let the UI decide)
        if best_match:
            result = {
                'label': label,
                **best_match
            }
        else:
            result = {
                'label': label,
                'category': '-',
                'type': 'income',
                'confidence': 0.0
            }

        self.cache[label] = result
        return result

    def categorize_batch(self, line_items):
        """Categorize multiple line items."""
        categorized = []
        for item in line_items:
            categorization = self.categorize_line_item(
                item['label'],
                sum(item['values']) if 'values' in item else 0
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
st.markdown("**Upload raw T12 → Auto-categorize → Review & Export**")

# ────────────────────────────────────────────────────────────────────────────
# UPLOAD T12
# ────────────────────────────────────────────────────────────────────────────

st.header("📁 Upload T12 Statement")

t12_file = st.file_uploader(
    "Upload raw T12 (Yardi/MRI format)",
    type=["xlsx", "xls"]
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
# CATEGORISATION
# ────────────────────────────────────────────────────────────────────────────

st.header("📋 Categorisation")

st.caption(f"**Property:** {parsed_t12['property_name']} | **Period:** {parsed_t12['period']}")

engine = CategorizationEngine()
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

# ────────────────────────────────────────────────────────────────────────────
# STEP 0: Select Total Income & NOI Line Items
# ────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("STEP 1: Mark Total Income & NOI Line Items")
st.markdown('**Select which line items represent "Total Income" and "NOI" to stop categorization after NOI**')

# Get list of all line items for dropdown
all_line_items = [row['line_item'] for row in table_data]

col_ti, col_noi = st.columns(2)

with col_ti:
    selected_total_income = st.selectbox(
        "Select Total Income Line Item:",
        options=['--'] + all_line_items,
        index=0,
        key="total_income_select"
    )

with col_noi:
    selected_noi = st.selectbox(
        "Select NOI Line Item:",
        options=['--'] + all_line_items,
        index=0,
        key="noi_select"
    )

# Store selections in session state
st.session_state.selected_total_income = selected_total_income
st.session_state.selected_noi = selected_noi

st.markdown("---")

if st.button("▶️ Proceed to Categorization", use_container_width=True, type="primary"):
    st.session_state.step1_complete = True
    st.rerun()

# Only show STEP 2 if STEP 1 is complete
if not st.session_state.get('step1_complete', False):
    st.stop()

# ────────────────────────────────────────────────────────────────────────────
# STEP 2: Categorisation
# ────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("STEP 2: Categorisation & Adjustments")
st.markdown("**Configure categories, multipliers, and amounts for each line item**")

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
    is_between_sections = False
    total_income_idx = None
    noi_idx = None

    # Find indices of Total Income and NOI
    for i, item in enumerate(table_data):
        if item['line_item'] == st.session_state.selected_total_income:
            total_income_idx = i
        if item['line_item'] == st.session_state.selected_noi:
            noi_idx = i

    if total_income_idx is not None and noi_idx is not None:
        is_between_sections = total_income_idx < idx < noi_idx

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
        category_options = ['-'] + list(engine.CATEGORY_RULES.keys())

        # Determine default category based on position
        if is_between_sections:
            # Between Total Income and NOI: default to Expense
            current_index = 0
            default_cat = "Expense"
        else:
            # Before Total Income or between NOI and end: use engine categorization
            should_be_empty = (original_amount == 0 and 'total' in row['line_item'].lower())

            if should_be_empty:
                current_index = 0  # Default to "-"
                default_cat = "-"
            elif row['category'] == '-':
                current_index = 0
                default_cat = "-"
            elif row['category'] in category_options:
                current_index = category_options.index(row['category'])
                default_cat = row['category']
            else:
                current_index = 0
                default_cat = "-"

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
st.header("✅ Financial Summary")

categorized_items = st.session_state.get('t12_categorized', [])

# Filter out post-NOI items for calculation
pre_noi_items = [i for i in categorized_items if not i.get('is_post_noi', False)]

# Calculate totals as per categorisation
total_income_categorized = sum([i['adjusted_amount'] for i in pre_noi_items if i['category'] == 'Gross Potential Rents'])

# Deductions (all the "Less:" categories)
deductions = sum([i['adjusted_amount'] for i in pre_noi_items if i['category'] in ['Less: Vacancy Loss', 'Less: Loss to Lease', 'Less: Non-Revenue Units', 'Less: Concessions', 'Less: Bad Debt']])

net_income = total_income_categorized - deductions

# Other Income
other_income = sum([i['adjusted_amount'] for i in pre_noi_items if i['category'] == 'Other Income'])

# Total Expenses
total_expenses = sum([i['adjusted_amount'] for i in pre_noi_items if i['category'] == 'Expense'])

# NOI = (Gross Potential Rents - Deductions + Other Income) - Expenses
noi_categorized = net_income + other_income - total_expenses

# Display As Per Categorisation
st.subheader("As Per Categorisation")
col1, col2, col3 = st.columns(3)

with col1:
    st.metric("💰 Total Income", f"${total_income_categorized:,.0f}")
with col2:
    st.metric("📉 Total Expenses", f"${total_expenses:,.0f}")
with col3:
    st.metric("📊 NOI", f"${noi_categorized:,.0f}")

# Display breakdown
st.markdown("---")
st.write("**Income Breakdown:**")
breakdown_cols = st.columns([2, 1])
with breakdown_cols[0]:
    st.write(f"Gross Potential Rents: ${total_income_categorized:,.0f}")
    st.write(f"Less: Deductions: -${deductions:,.0f}")
    st.write(f"Net Income: ${net_income:,.0f}")
    st.write(f"Plus: Other Income: ${other_income:,.0f}")
    st.divider()
    st.write(f"**Total Rental Income: ${net_income + other_income:,.0f}**")

st.write("**Expense Breakdown:**")
with breakdown_cols[0]:
    st.write(f"Total Operating Expenses: ${total_expenses:,.0f}")

st.divider()
st.write(f"**Net Operating Income (NOI): ${noi_categorized:,.0f}**")

if st.session_state.selected_noi != '--':
    st.info(f"✅ Categorization stops after: **{st.session_state.selected_noi}**")

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
