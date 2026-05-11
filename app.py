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
                continue

            label = str(col_a).strip()
            if not label or len(label) < 2:
                continue

            # Skip header/summary rows
            if any(x in label.lower() for x in ['total', 'summary', 'header', 'property occupancy']):
                continue

            # Collect monthly values
            values = []
            for c in range(2, min(14, self.ws.max_column + 1)):
                try:
                    val = self.ws.cell(r, c).value
                    if val is not None:
                        values.append(float(val) if isinstance(val, (int, float)) else 0)
                    else:
                        values.append(0)
                except:
                    values.append(0)

            # Only include rows with data
            if any(v != 0 for v in values):
                line_items.append({
                    'label': label,
                    'values': values,
                    'is_subtotal': 'total' in label.lower(),
                    'is_section_header': False
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

# Prepare table data
table_data = []
item_idx = 0

for item in categorized:
    if item['is_subtotal'] or item['is_section_header']:
        continue

    total_val = sum(item['values']) if 'values' in item else 0

    # Detect multiplier
    multiplier_text = "×1"
    if total_val != 0:
        if ('utility' in item['label'].lower() or 'rubs' in item['label'].lower()) and total_val < 0:
            multiplier_text = "×-1"

    table_data.append({
        'idx': item_idx,
        'amount': total_val,
        'line_item': item['label'],
        'category': item['category'],
        'income_expense_type': 'Income' if item['type'] == 'income' else 'Expense',
        'multiplier': multiplier_text
    })
    item_idx += 1

# Mark where Income ends / NOI begins
st.markdown("**Select where Income ends and NOI begins:**")

col_marker, col_blank = st.columns([3, 2])
with col_marker:
    income_end_idx = st.selectbox(
        "Income Ends After:",
        options=[f"Item {i+1}: {row['line_item'][:40]}" for i, row in enumerate(table_data)],
        index=0,
        help="Select the last item that is part of Gross Income. Everything after is NOI."
    )

# Extract the selected index
income_end_position = int(income_end_idx.split(":")[0].replace("Item ", "")) - 1

st.markdown("---")
st.markdown("**Line Item | Amount | Category | Multiplier**")

edited_items = []
for idx, row in enumerate(table_data):
    # Determine if this is Income or NOI
    if idx <= income_end_position:
        section_type = "Income"
        section_marker = "💰"
    else:
        section_type = "NOI"
        section_marker = "📉"

    col1, col2, col3, col4, col5 = st.columns([0.5, 2.5, 1.2, 1.8, 0.8])

    with col1:
        st.write(section_marker)

    with col2:
        st.write(f"`{row['line_item'][:40]}`")

    with col3:
        st.write(f"**{row['amount']:,.0f}**")

    with col4:
        selected_cat = st.selectbox(
            "Category",
            options=list(engine.CATEGORY_RULES.keys()),
            index=list(engine.CATEGORY_RULES.keys()).index(row['category']) if row['category'] in engine.CATEGORY_RULES else 0,
            key=f"cat_{idx}_{hash(row['line_item']) % 10000}",
            label_visibility="collapsed"
        )

    with col5:
        st.write(f"`{row['multiplier']}`")

    edited_items.append({
        'label': row['line_item'],
        'amount': row['amount'],
        'category': selected_cat,
        'section_type': section_type,
        'multiplier': row['multiplier'],
        'values': table_data[idx]
    })

st.session_state.t12_categorized = edited_items

st.markdown("---")
st.header("✅ Summary")

col_a, col_b, col_c = st.columns(3)
with col_a:
    st.metric("Total Line Items", len(edited_items))
with col_b:
    income_count = len([i for i in edited_items if i['section_type'] == 'Income'])
    st.metric("💰 Income Items", income_count)
with col_c:
    noi_count = len([i for i in edited_items if i['section_type'] == 'NOI'])
    st.metric("📉 NOI Items", noi_count)

st.markdown("---")
st.subheader("📥 Export Categorized T12")

if st.button("✅ Export as CSV", use_container_width=True, type="primary"):
    export_data = []
    for item in st.session_state.get('t12_categorized', []):
        export_data.append({
            'Line Item': item['label'],
            'Amount': item['amount'],
            'Category': item['category'],
            'Section': item['section_type'],
            'Multiplier': item['multiplier']
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
