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
        for r in range(1, min(20, self.ws.max_row + 1)):
            cell_val = self.ws.cell(r, 1).value
            if cell_val and isinstance(cell_val, str):
                if not property_name:
                    property_name = str(cell_val).strip()
                if 'month' in str(cell_val).lower() or 'period' in str(cell_val).lower():
                    period = str(cell_val).strip()

        # Find data rows (scan for numeric data)
        for r in range(1, self.ws.max_row + 1):
            col_a = self.ws.cell(r, 1).value
            if col_a is None:
                col_a = self.ws.cell(r, 2).value

            if col_a is None:
                continue

            label = str(col_a).strip()
            if not label or len(label) < 1:
                continue

            # Skip only pure summary rows (not section headers)
            if 'property occupancy' in label.lower():
                continue

            # Detect section headers (low indent, typically uppercase or key section terms)
            is_section = any(x in label.upper() for x in ['GROSS', 'LESS:', 'EXPENSE', 'INCOME', 'OPERATION', 'UTILITIES', 'PAYROLL'])
            is_subtotal = 'TOTAL' in label.upper()

            # Collect monthly values - scan all columns with numeric data
            values = []
            has_any_value = False
            for c in range(2, min(20, self.ws.max_column + 1)):
                try:
                    val = self.ws.cell(r, c).value
                    if val is not None and isinstance(val, (int, float)):
                        values.append(float(val))
                        if val != 0:
                            has_any_value = True
                    else:
                        values.append(0)
                except:
                    values.append(0)

            # Include rows with label + either has data or is a section header
            if has_any_value or is_section or is_subtotal:
                # Pad to 12 months if needed
                while len(values) < 12:
                    values.append(0)

                line_items.append({
                    'label': label,
                    'values': values[:12],  # Take first 12 months
                    'is_subtotal': is_subtotal,
                    'is_section_header': is_section
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
            'patterns': [r'\bGross\s+Potential\s+Rent\b', r'\bGPR\b', r'^\s*Gross Rental Income$'],
            'type': 'income'
        },
        'Less: Vacancy Loss': {
            'patterns': [r'Vacancy\s+Loss', r'Vacant Unit', r'Unoccupied'],
            'type': 'income'
        },
        'Less: Loss to Lease': {
            'patterns': [r'Loss.*to.*Lease', r'Lease.*Loss'],
            'type': 'income'
        },
        'Less: Non-Revenue Units': {
            'patterns': [r'Non[\-]?Revenue', r'Model', r'Admin'],
            'type': 'income'
        },
        'Less: Concessions': {
            'patterns': [r'Rent\s+Concessions?', r'Concession', r'Lease Concession'],
            'type': 'income'
        },
        'Less: Bad Debt': {
            'patterns': [r'Bad\s+Debt', r'Allowance.*Doubtful'],
            'type': 'income'
        },
        'Other Income': {
            'patterns': [r'Other\s+Income', r'Pet\s+Fees', r'Parking', r'Laundry', r'Late\s+Fees', r'RUBS'],
            'type': 'income'
        },
        'Expense': {
            'patterns': [r'Expense', r'Payroll', r'Utilities', r'Repairs', r'Management', r'Insurance', r'Taxes'],
            'type': 'expense'
        }
    }

    def __init__(self):
        self.cache = {}

    def categorize_line_item(self, label, value=0):
        """Categorize a single line item."""
        if label in self.cache:
            return self.cache[label]

        result = {
            'label': label,
            'category': 'Other Income',
            'type': 'income',
            'confidence': 0.0
        }

        best_match = None
        best_score = 0

        for category, rules in self.CATEGORY_RULES.items():
            for pattern in rules['patterns']:
                match = re.search(pattern, label, re.IGNORECASE)
                if match:
                    score = 100 - (match.start() / len(label) * 20)
                    if score > best_score:
                        best_score = score
                        best_match = {
                            'category': category,
                            'type': rules['type'],
                            'confidence': min(1.0, score / 100)
                        }

        if best_match:
            result.update(best_match)

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

# Prepare table data (include section headers)
table_data = []
item_idx = 0

for item in categorized:
    if item['is_subtotal']:
        continue

    total_val = sum(item['values']) if 'values' in item else 0

    # Detect multiplier
    multiplier_text = "×1"
    if total_val != 0:
        if ('utility' in item['label'].lower() or 'rubs' in item['label'].lower()) and total_val < 0:
            multiplier_text = "×-1"

    is_section = item['is_section_header']

    table_data.append({
        'idx': item_idx,
        'amount': total_val,
        'line_item': item['label'],
        'category': item['category'],
        'income_expense_type': 'Income' if item['type'] == 'income' else 'Expense',
        'multiplier': multiplier_text,
        'is_section': is_section
    })
    item_idx += 1

# STEP 1: Mark Income/NOI sections
st.markdown("---")
st.subheader("STEP 1: Mark Income/NOI Sections")
st.markdown('**Select "-", "Total Income", or "NOI" for each line item**')

# Table header for section marking
header_col1, header_col2, header_col3, header_col4, header_col5, header_col6 = st.columns([0.8, 2.2, 1.0, 1.5, 1.5, 0.8])
with header_col1:
    st.markdown("**Income/NOI**")
with header_col2:
    st.markdown("**Line Item**")
with header_col3:
    st.markdown("**Amount**")
with header_col4:
    st.markdown("**Category**")
with header_col5:
    st.markdown("**Type**")
with header_col6:
    st.markdown("**Mult.**")

st.markdown("---")

# Table rows - Step 1: Mark sections only
section_selections = []
for idx, row in enumerate(table_data):
    col1, col2, col3, col4, col5, col6 = st.columns([0.8, 2.2, 1.0, 1.5, 1.5, 0.8])

    # Style section headers differently
    if row['is_section']:
        with col1:
            st.write("📌")  # Section header marker
        with col2:
            st.markdown(f"**{row['line_item'].upper()}**")  # Bold, uppercase
        with col3:
            st.write("")
        with col4:
            st.write("")
        with col5:
            st.write("")
        with col6:
            st.write("")

        section_selections.append({
            'section': '-',
            'data': row,
            'is_section_header': True
        })
    else:
        with col1:
            selected_section = st.selectbox(
                "Section",
                options=['-', 'Total Income', 'NOI'],
                index=0,
                key=f"section_{idx}_{hash(row['line_item']) % 10000}",
                label_visibility="collapsed"
            )

        with col2:
            st.write(f"`{row['line_item'][:35]}`")

        with col3:
            st.write(f"**{row['amount']:,.0f}**")

        with col4:
            st.write(f"`{row['category']}`")

        with col5:
            st.write(f"`{row['income_expense_type']}`")

        with col6:
            st.write(f"`{row['multiplier']}`")

        section_selections.append({
            'section': selected_section,
            'data': row,
            'is_section_header': False
        })

# Store section selections and show "Run Categorisation" button
st.markdown("---")

if st.button("▶️ Run Categorisation", use_container_width=True, type="primary"):
    st.session_state.sections_marked = True
    st.session_state.section_selections = section_selections
    st.rerun()

# STEP 2: Categorisation (only show if sections are marked)
if st.session_state.get('sections_marked', False):
    st.markdown("---")
    st.subheader("STEP 2: Categorisation")
    st.markdown("**Select categories for each line item**")

    # Table header for categorisation
    header_col1, header_col2, header_col3, header_col4, header_col5, header_col6, header_col7 = st.columns([0.9, 2.0, 1.0, 0.7, 1.0, 2.0, 0.6])
    with header_col1:
        st.markdown("**Income/NOI**")
    with header_col2:
        st.markdown("**Line Item**")
    with header_col3:
        st.markdown("**Orig. Amt**")
    with header_col4:
        st.markdown("**Mult.**")
    with header_col5:
        st.markdown("**Adj. Amt**")
    with header_col6:
        st.markdown("**Category**")
    with header_col7:
        st.markdown("**Marker**")

    st.markdown("---")

    # Table rows - Step 2: Categorisation
    edited_items = []
    for idx, selection in enumerate(st.session_state.section_selections):
        row = selection['data']
        selected_section = selection['section']
        original_amount = row['amount']
        is_section_header = selection.get('is_section_header', False)

        col1, col2, col3, col4, col5, col6, col7 = st.columns([0.9, 2.0, 1.0, 0.7, 1.0, 2.0, 0.6])

        # Section headers have different styling
        if is_section_header:
            with col1:
                st.write("📌")
            with col2:
                st.markdown(f"**{row['line_item'].upper()}**")
            with col3:
                st.write("")
            with col4:
                st.write("")
            with col5:
                st.write("")
            with col6:
                st.write("")
            with col7:
                st.write("")

            edited_items.append({
                'label': row['line_item'],
                'original_amount': original_amount,
                'multiplier': 1,
                'adjusted_amount': original_amount,
                'category': '-',
                'section_type': '-',
                'values': row,
                'is_section_header': True
            })
        else:
            with col1:
                st.write(f"**{selected_section}**")

            with col2:
                st.write(f"`{row['line_item'][:35]}`")

            with col3:
                st.write(f"**{original_amount:,.0f}**")

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
                st.write(f"**{adjusted_amount:,.0f}**")

            with col6:
                selected_cat = st.selectbox(
                    "Category",
                    options=list(engine.CATEGORY_RULES.keys()),
                    index=list(engine.CATEGORY_RULES.keys()).index(row['category']) if row['category'] in engine.CATEGORY_RULES else 0,
                    key=f"cat_{idx}_{hash(row['line_item']) % 10000}",
                    label_visibility="collapsed"
                )

            with col7:
                marker_text = "🔄" if multiplier_value == -1 else ""
                st.write(marker_text)

            edited_items.append({
                'label': row['line_item'],
                'original_amount': original_amount,
                'multiplier': multiplier_value,
                'adjusted_amount': adjusted_amount,
                'category': selected_cat,
                'section_type': selected_section,
                'values': row,
                'is_section_header': False
            })

    st.session_state.t12_categorized = edited_items

if st.session_state.get('sections_marked', False):
    st.markdown("---")
    st.header("✅ Financial Summary")

    categorized_items = st.session_state.get('t12_categorized', [])

    # Calculate totals as per categorisation
    total_income_categorized = sum([i['adjusted_amount'] for i in categorized_items if i['category'] == 'Gross Potential Rents'])

    # Deductions (all the "Less:" categories)
    deductions = sum([i['adjusted_amount'] for i in categorized_items if i['category'] in ['Less: Vacancy Loss', 'Less: Loss to Lease', 'Less: Non-Revenue Units', 'Less: Concessions', 'Less: Bad Debt']])

    net_income = total_income_categorized - deductions

    # Other Income
    other_income = sum([i['adjusted_amount'] for i in categorized_items if i['category'] == 'Other Income'])

    # Total Expenses
    total_expenses = sum([i['adjusted_amount'] for i in categorized_items if i['category'] == 'Expense'])

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

    st.markdown("---")
    st.subheader("📥 Export Categorized T12")

    if st.button("✅ Export as CSV", use_container_width=True, type="primary"):
        export_data = []
        for item in st.session_state.get('t12_categorized', []):
            export_data.append({
                'Line Item': item['label'],
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
