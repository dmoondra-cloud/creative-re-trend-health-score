"""
Creative RE — T12 Categorization Tool
======================================
Simple T12 line-item parser and categorizer.
Upload raw T12 → Auto-categorize → Review & Edit → Export
"""

import io
import os
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
from pathlib import Path
import tempfile
import traceback

from t12_parser import T12Parser
from categorization_engine import CategorizationEngine

# ============================================================================
# PAGE CONFIG
# ============================================================================

st.set_page_config(
    page_title="Creative RE T12 Categorizer",
    page_icon="📊",
    layout="wide"
)

st.title("📊 T12 Categorization Tool")
st.markdown("**Upload raw T12 → Auto-categorize → Review & Export**")

# ============================================================================
# UPLOAD T12
# ============================================================================

st.header("📁 Upload T12 Statement")

t12_file = st.file_uploader(
    "Upload raw T12 (Yardi/MRI format)",
    type=["xlsx", "xls"],
    key="t12_upload"
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
        st.error("❌ Could not parse T12. Check file format.")
        st.stop()

except Exception as e:
    st.error(f"❌ Error parsing T12: {str(e)}")
    st.stop()

# ============================================================================
# CATEGORISATION
# ============================================================================

st.header("📋 Categorisation")

st.caption(f"**Property:** {parsed_t12['property_name']} | **Period:** {parsed_t12.get('period', 'N/A')}")

# Categorize T12
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
header_col1, header_col2, header_col3, header_col4 = st.columns([1.2, 3.0, 1.5, 1.0])
with header_col1:
    st.markdown("**Income/NOI**")
with header_col2:
    st.markdown("**Line Item**")
with header_col3:
    st.markdown("**Amount**")
with header_col4:
    st.markdown("**Multiplier**")

st.markdown("---")

# Table rows - Step 1: Mark sections only
section_selections = []
for idx, row in enumerate(table_data):
    col1, col2, col3, col4 = st.columns([1.2, 3.0, 1.5, 1.0])

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
            st.write(f"`{row['line_item'][:40]}`")

        with col3:
            st.write(f"**{row['amount']:,.0f}**")

        with col4:
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
    header_col1, header_col2, header_col3, header_col4, header_col5, header_col6, header_col7 = st.columns([1.0, 2.2, 1.1, 1.0, 1.2, 1.3, 0.7])
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

        col1, col2, col3, col4, col5, col6, col7 = st.columns([1.0, 2.2, 1.1, 1.0, 1.2, 1.3, 0.7])

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

# ============================================================================
# SUMMARY & EXPORT (Only show if categorisation is done)
# ============================================================================

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
        # Create DataFrame from categorized items
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
            label="📥 Download Categorized T12 (CSV)",
            data=csv,
            file_name=f"T12_Categorized_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
        st.success("✅ Ready to download!")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center'>
    <p><small>Creative RE T12 Categorization Tool | Auto-detect & Edit Categories</small></p>
</div>
""", unsafe_allow_html=True)
