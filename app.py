"""
Creative RE — UW Pipeline (v2)
==============================
Rebuilt with:
- Pre-loaded THS template
- Property details form
- Dual T12/RR processing sections
- Real-time T12 categorization with edit option
- RR parsing using exact old script logic
- THS-only output (no separate T12)
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
from rr_processor import RRProcessor
from ths_processor import THSProcessor

# ============================================================================
# PAGE CONFIG
# ============================================================================

st.set_page_config(
    page_title="Creative RE Underwriting Suite",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Creative RE Underwriting Suite")
st.markdown("**T12 + RR Processing → Trend Health Score**")

# ============================================================================
# LOAD THS TEMPLATE (PRE-LOADED)
# ============================================================================

@st.cache_resource
def load_ths_template():
    """Load pre-loaded THS template from templates folder."""
    # Try multiple path variations for Streamlit Cloud compatibility
    paths_to_try = [
        "templates/THS_Template_Default.xlsx",
        "./templates/THS_Template_Default.xlsx",
        os.path.join(os.getcwd(), "templates", "THS_Template_Default.xlsx"),
        os.path.join(os.path.dirname(__file__), "templates", "THS_Template_Default.xlsx"),
    ]

    for template_path in paths_to_try:
        if os.path.exists(template_path):
            try:
                return openpyxl.load_workbook(template_path)
            except Exception as e:
                st.error(f"Error loading template from {template_path}: {str(e)}")
                return None

    # If no path worked, show error with debugging info
    st.error(f"❌ Template not found in any of these locations:")
    for path in paths_to_try:
        st.error(f"  - {path}")
    st.error(f"Current working directory: {os.getcwd()}")
    return None

ths_template = load_ths_template()
if ths_template is None:
    st.error("❌ THS Template not found. Please add to templates/ folder.")
    st.stop()

# ============================================================================
# STEP 1: PROPERTY DETAILS FORM
# ============================================================================

st.header("📋 STEP 1: Property Information")
col1, col2 = st.columns(2)

with col1:
    property_name = st.text_input(
        "Property Name *",
        placeholder="e.g., Woodward Park"
    )

with col2:
    property_address = st.text_input(
        "Property Address *",
        placeholder="e.g., 111 NW 63rd Ave, Bentonville, AR 72713"
    )

property_url = st.text_input(
    "Apartments.com URL (optional)",
    placeholder="https://www.apartments.com/..."
)

if not property_name or not property_address:
    st.warning("⚠️ Please enter property name and address to continue")
    st.stop()

# ============================================================================
# STEP 2: FILE UPLOADS & RR DETAILS
# ============================================================================

st.header("📁 STEP 2: Upload Files & RR Details")

col_t12, col_rr = st.columns(2)

with col_t12:
    st.subheader("T12 Statement (Raw)")
    t12_file = st.file_uploader(
        "Upload raw T12 (Yardi/MRI format)",
        type=["xlsx", "xls"],
        key="t12_upload"
    )

with col_rr:
    st.subheader("Rent Roll (Raw)")
    rr_file = st.file_uploader(
        "Upload raw Rent Roll",
        type=["xlsx", "xls"],
        key="rr_upload"
    )

# RR Details
st.subheader("Rent Roll Details")
col_rr_date, col_floorplans = st.columns(2)

with col_rr_date:
    rr_date = st.date_input(
        "RR As-Of Date *",
        value=datetime.now()
    )

with col_floorplans:
    st.markdown("**Floor Plan Sizes** (optional)")
    floor_plans_text = st.text_area(
        "Enter one per line: '2 BR: 904' or '3 BR: 1101, 1194'",
        height=100,
        placeholder="2 BR: 904\n3 BR: 1101, 1194\n4 BR: 1327"
    )

if not t12_file or not rr_file:
    st.warning("⚠️ Please upload both T12 and Rent Roll files to continue")
    st.stop()

# ============================================================================
# STEP 3: PARSE FILES
# ============================================================================

st.header("⚙️ STEP 3: Process & Review")

# Parse T12
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

# Parse RR
try:
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(rr_file.getbuffer())
        rr_path = tmp.name

    rr_processor = RRProcessor(rr_path)
    rr_data = rr_processor.load_rr()
    rr_summary = rr_processor.get_summary()

except Exception as e:
    st.error(f"❌ Error processing RR: {str(e)}")
    st.stop()

# ============================================================================
# STEP 4: DUAL PROCESSING SECTION
# ============================================================================

st.markdown("---")
st.subheader("📊 Processing Results")

col_t12_section, col_rr_section = st.columns(2)

# ══════════════════════════════════════════════════════════════════════════
# T12 SECTION (LEFT)
# ══════════════════════════════════════════════════════════════════════════

with col_t12_section:
    st.markdown("### 📋 T12 Categorization")
    st.caption(f"Property: **{parsed_t12['property_name']}**")

    # Categorize T12
    engine = CategorizationEngine()
    categorized = engine.categorize_batch(parsed_t12['line_items'])

    # Display categorization with edit option
    st.markdown("**Review & Edit Categories:**")

    # Store categorized items for editing
    if 'edited_categories' not in st.session_state:
        st.session_state.edited_categories = {}

    edited_items = []

    for item in categorized:
        if item['is_subtotal'] or item['is_section_header']:
            continue

        # Calculate total value
        total_val = sum(item['values']) if 'values' in item else 0

        # Check if value was multiplied by -1
        multiplier = ""
        if total_val != 0:
            # Detect if this is RUBS/utility that was negated
            if ('utility' in item['label'].lower() or 'rubs' in item['label'].lower()) and total_val < 0:
                multiplier = " **[×-1]**"

        # Create editable entry
        col_label, col_category, col_value = st.columns([2, 1.5, 1])

        with col_label:
            st.markdown(f"`{item['label'][:40]}`")

        with col_category:
            current_cat = item['category']
            st.selectbox(
                "Category",
                options=list(engine.CATEGORY_RULES.keys()),
                index=list(engine.CATEGORY_RULES.keys()).index(current_cat) if current_cat in engine.CATEGORY_RULES else 0,
                key=f"cat_{item['label']}",
                label_visibility="collapsed"
            )

        with col_value:
            st.markdown(f"**{total_val:,.0f}**{multiplier}")

        edited_items.append({
            'label': item['label'],
            'category': st.session_state.get(f"cat_{item['label']}", item['category']),
            'values': item['values'],
            'original': item
        })

    st.session_state.t12_categorized = edited_items

    # T12 Summary
    st.markdown("**Categorization Summary:**")
    col_t12_a, col_t12_b, col_t12_c = st.columns(3)
    with col_t12_a:
        st.metric("Line Items", len([i for i in categorized if not i['is_section_header']]))
    with col_t12_b:
        st.metric("Income Categories", len(set(i['category'] for i in categorized if i['type'] == 'income')))
    with col_t12_c:
        st.metric("Expense Categories", len(set(i['category'] for i in categorized if i['type'] == 'expense')))

# ══════════════════════════════════════════════════════════════════════════
# RR SECTION (RIGHT)
# ══════════════════════════════════════════════════════════════════════════

with col_rr_section:
    st.markdown("### 📋 Rent Roll Summary")
    st.caption(f"As-Of Date: **{rr_date.strftime('%m/%d/%Y')}**")

    # RR Summary Stats
    col_rr_a, col_rr_b, col_rr_c = st.columns(3)
    with col_rr_a:
        st.metric("Total Units", rr_summary['total_units'])

    # Try to get occupancy
    if 'occupancy_stats' in rr_summary:
        rented = rr_summary['occupancy_stats'].get('Rented', 0)
        with col_rr_b:
            st.metric("Rented", rented)
        with col_rr_c:
            vacant = rr_summary['total_units'] - rented
            st.metric("Vacant", vacant)

    # RR Data Preview
    st.markdown("**Data Preview:**")
    st.dataframe(
        rr_summary['data'].head(10),
        use_container_width=True,
        height=200
    )

    st.session_state.rr_data = rr_summary['data']

# ============================================================================
# STEP 5: APPROVE & GENERATE THS
# ============================================================================

st.markdown("---")

col_approve, col_status = st.columns([1, 3])

with col_approve:
    if st.button("✅ Approve & Generate THS", use_container_width=True, type="primary"):
        st.session_state.generate_ths = True

# If approved, generate THS
if st.session_state.get('generate_ths', False):
    st.markdown("---")
    st.subheader("📊 Generating Trend Health Score...")

    try:
        # Populate T12 into template
        from t12_template_processor import T12TemplateProcessor

        processor = T12TemplateProcessor(None)  # Will use loaded template
        processor.ws = ths_template['T12']
        processor.wb = ths_template

        # Update metadata
        processor.ws['C4'] = property_name
        processor.ws['C6'] = rr_date.strftime('%m/%d/%Y')

        # Populate categorized items
        for idx, item in enumerate(st.session_state.get('t12_categorized', []), start=9):
            processor.ws.cell(idx, 4).value = item['label']  # Particulars
            processor.ws.cell(idx, 3).value = item['category']  # Category

            for col_idx, val in enumerate(item['values'][:12], start=5):
                processor.ws.cell(idx, col_idx).value = val

        # Save to bytes
        output = io.BytesIO()
        ths_template.save(output)
        output.seek(0)

        # Offer download
        st.download_button(
            label="📥 Download THS Template",
            data=output,
            file_name=f"{property_name}_THS_{rr_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ THS generated successfully!")
        st.balloons()

    except Exception as e:
        st.error(f"❌ Error generating THS: {str(e)}")
        with st.expander("View error details"):
            st.code(traceback.format_exc())

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center'>
    <p><small>Creative RE Underwriting Suite | Format-Agnostic T12 Parser</small></p>
</div>
""", unsafe_allow_html=True)
