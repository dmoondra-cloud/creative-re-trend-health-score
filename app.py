"""
Creative RE T12 + THS Underwriting App
Format-agnostic T12 parser with automated categorization and THS processing.
Built with Streamlit for easy deployment.
"""

import streamlit as st
import pandas as pd
import openpyxl
from pathlib import Path
import tempfile
import io
from typing import Optional
import traceback

from t12_parser import T12Parser
from categorization_engine import CategorizationEngine
from t12_template_processor import T12TemplateProcessor, T12Validator
from rr_processor import RRProcessor
from ths_processor import THSProcessor


# ============================================================================
# PAGE CONFIG
# ============================================================================

st.set_page_config(
    page_title="Creative RE Underwriting Suite",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# SIDEBAR CONFIG & FILE UPLOAD
# ============================================================================

st.sidebar.title("📋 Creative RE UW Suite")
st.sidebar.markdown("---")

# Upload sections
st.sidebar.subheader("1. Upload Raw T12")
t12_file = st.sidebar.file_uploader(
    "Upload raw T12 statement (Yardi/MRI format)",
    type=["xlsx", "xls"],
    key="t12_upload"
)

st.sidebar.subheader("2. Upload Rent Roll")
rr_file = st.sidebar.file_uploader(
    "Upload Rent Roll",
    type=["xlsx", "xls"],
    key="rr_upload"
)

st.sidebar.subheader("3. Select THS Template")
template_file = st.sidebar.file_uploader(
    "Upload THS Template (or use default)",
    type=["xlsx"],
    key="template_upload"
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Need help?** Check out the documentation tab →")

# ============================================================================
# MAIN PAGE
# ============================================================================

tabs = st.tabs([
    "📊 Dashboard",
    "🔄 T12 Categorization",
    "📋 Rent Roll",
    "✅ Review & Export",
    "📖 Help"
])

# ============================================================================
# TAB 1: DASHBOARD
# ============================================================================

with tabs[0]:
    st.title("📊 Underwriting Dashboard")
    st.markdown("""
    Welcome to Creative RE's automated underwriting suite. This app helps you:
    1. **Parse** raw T12 statements in any format (Yardi, MRI, etc.)
    2. **Auto-categorize** line items with intelligent pattern matching
    3. **Process** Rent Rolls
    4. **Generate** Trend Health Score reports
    5. **Export** all workbooks for further analysis
    """)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("T12 Status", "✓ Ready" if t12_file else "⏳ Pending")

    with col2:
        st.metric("RR Status", "✓ Ready" if rr_file else "⏳ Pending")

    with col3:
        st.metric("Template Status", "✓ Ready" if template_file else "ℹ️ Using default")

    st.markdown("---")
    st.subheader("📝 Next Steps:")
    if not t12_file:
        st.info("👈 Upload a T12 statement to begin")
    elif not rr_file:
        st.info("👈 Upload a Rent Roll to continue")
    else:
        st.success("✅ All files ready! Move to 'T12 Categorization' tab →")


# ============================================================================
# TAB 2: T12 CATEGORIZATION
# ============================================================================

with tabs[1]:
    st.title("🔄 T12 Categorization")

    if not t12_file:
        st.warning("Please upload a raw T12 file first using the sidebar ←")
    else:
        st.subheader("Parsing T12 Statement...")

        try:
            # Save uploaded file to temp location
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(t12_file.getbuffer())
                tmp_path = tmp.name

            # Parse the file
            parser = T12Parser(tmp_path)
            parsed_data = parser.parse()

            if not parsed_data['parsed_successfully']:
                st.error("Could not parse T12 file. Is it in the correct format?")
            else:
                # Display extracted metadata
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"**Property:** {parsed_data['property_name']}")
                with col2:
                    st.info(f"**As-Of Date:** {parsed_data['as_of_date'] or 'Not detected'}")

                st.markdown("---")
                st.subheader("📊 Line Items Detected")

                # Display line items
                line_items_df = pd.DataFrame([
                    {
                        'Label': item['label'],
                        'Indent': item['indent'],
                        'Is Subtotal': item['is_subtotal'],
                        'Values': f"{len(item['values'])} months"
                    }
                    for item in parsed_data['line_items']
                ])

                st.dataframe(line_items_df, use_container_width=True)

                st.markdown("---")
                st.subheader("🏷️ Auto-Categorizing Line Items...")

                # Categorize
                engine = CategorizationEngine()
                categorized = engine.categorize_batch(parsed_data['line_items'])

                # Display categorization results
                cat_df = pd.DataFrame([
                    {
                        'Line Item': item['label'][:40] + ('...' if len(item['label']) > 40 else ''),
                        'Category': item['category'],
                        'Type': item['type'],
                        'Confidence': f"{item.get('confidence', 0)*100:.0f}%"
                    }
                    for item in categorized
                    if not item['is_subtotal'] and not item['is_section_header']
                ])

                st.dataframe(cat_df, use_container_width=True)

                # Show any low-confidence items
                low_conf = [
                    item for item in categorized
                    if item.get('confidence', 0) < 0.7 and not item['is_section_header']
                ]

                if low_conf:
                    st.warning(f"⚠️ {len(low_conf)} items with low categorization confidence")
                    with st.expander("View low-confidence items"):
                        for item in low_conf:
                            st.write(f"- **{item['label']}** → {item['category']} ({item.get('confidence', 0)*100:.0f}%)")

                # Store in session state for next tab
                st.session_state.parsed_data = parsed_data
                st.session_state.categorized_data = categorized

                st.success("✅ T12 categorization complete! Move to 'Review & Export' tab →")

        except Exception as e:
            st.error(f"Error parsing T12: {str(e)}")
            with st.expander("View error details"):
                st.code(traceback.format_exc())


# ============================================================================
# TAB 3: RENT ROLL
# ============================================================================

with tabs[2]:
    st.title("📋 Rent Roll Processing")

    if not rr_file:
        st.warning("Please upload a Rent Roll file using the sidebar ←")
    else:
        st.subheader("Processing Rent Roll...")

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(rr_file.getbuffer())
                tmp_path = tmp.name

            # Load and display RR stats
            rr_df = pd.read_excel(tmp_path, sheet_name='Rent Roll' if 'Rent Roll' in
                                 pd.ExcelFile(tmp_path).sheet_names else 0)

            st.subheader("📊 Rent Roll Summary")
            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("Total Units", len(rr_df))

            with col2:
                try:
                    rented = len(rr_df[rr_df['Status'] == 'Rented'])
                    st.metric("Rented Units", rented)
                except:
                    st.metric("Rented Units", "N/A")

            with col3:
                try:
                    vacant = len(rr_df[rr_df['Status'] == 'Vacant'])
                    st.metric("Vacant Units", vacant)
                except:
                    st.metric("Vacant Units", "N/A")

            st.markdown("---")
            st.subheader("Rent Roll Data")
            st.dataframe(rr_df.head(20), use_container_width=True)

            st.session_state.rr_data = rr_df

        except Exception as e:
            st.error(f"Error processing Rent Roll: {str(e)}")


# ============================================================================
# TAB 4: REVIEW & EXPORT
# ============================================================================

with tabs[3]:
    st.title("✅ Review & Export")

    if 'categorized_data' not in st.session_state:
        st.warning("Please complete T12 categorization first ←")
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📋 Export Options")

            if st.button("📥 Generate T12 Template", use_container_width=True):
                try:
                    if not template_file:
                        st.info("Creating default T12 template...")
                        st.warning("No default template bundled. Please provide template file.")
                    else:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                            tmp.write(template_file.getbuffer())
                            template_path = tmp.name

                        # Process and populate T12
                        processor = T12TemplateProcessor(template_path)
                        populated_wb = processor.populate_t12(
                            st.session_state.categorized_data,
                            st.session_state.parsed_data['property_name'],
                            st.session_state.parsed_data['as_of_date'],
                            st.session_state.parsed_data['months']
                        )

                        # Save to bytes
                        output = io.BytesIO()
                        populated_wb.save(output)
                        output.seek(0)

                        st.session_state.t12_output = output
                        st.success("✅ T12 template populated!")

                except Exception as e:
                    st.error(f"Error generating T12: {str(e)}")

            if st.button("📥 Generate THS Report", use_container_width=True):
                st.info("THS generation requires additional metric inputs...")

        with col2:
            st.subheader("💾 Download Files")

            if 't12_output' in st.session_state:
                st.download_button(
                    label="📥 Download T12 Template",
                    data=st.session_state.t12_output,
                    file_name=f"{st.session_state.parsed_data['property_name']}_T12.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        st.markdown("---")
        st.subheader("📊 Categorization Summary")

        # Summary stats
        categorized = st.session_state.categorized_data
        income_items = [item for item in categorized if item['type'] == 'income']
        expense_items = [item for item in categorized if item['type'] == 'expense']

        summary_col1, summary_col2, summary_col3 = st.columns(3)
        with summary_col1:
            st.metric("Total Line Items", len(categorized))
        with summary_col2:
            st.metric("Income Categories", len(set(i['category'] for i in income_items)))
        with summary_col3:
            st.metric("Expense Categories", len(set(i['category'] for i in expense_items)))


# ============================================================================
# TAB 5: HELP
# ============================================================================

with tabs[4]:
    st.title("📖 Help & Documentation")

    st.subheader("🚀 Quick Start")
    st.markdown("""
    1. **Upload Raw T12** - Any Yardi/MRI format statement
    2. **Review Categorization** - Check auto-categorized line items
    3. **Upload Rent Roll** - Include latest RR
    4. **Generate Reports** - Create T12 and THS documents
    5. **Download** - Export all workbooks

    **The app automatically:**
    - ✅ Detects property name and as-of-date
    - ✅ Categorizes GPR, concessions, rental losses
    - ✅ Handles RUBS, utility income, reimbursements
    - ✅ Converts negative expenses to income when appropriate
    - ✅ Preserves cell validations
    """)

    st.subheader("❓ FAQ")
    with st.expander("What formats does it support?"):
        st.write("""
        The app handles multiple T12 formats including Yardi and MRI statements.
        As long as there's a clear line item structure with months, it works.
        """)

    with st.expander("How does auto-categorization work?"):
        st.write("""
        Pattern matching on line item descriptions:
        - Keywords (GPR, concessions, utilities, RUBS)
        - Matches to Creative RE standard categories
        - Flags low-confidence items for review
        """)

    st.markdown("---")
    st.subheader("📧 Support")
    st.write("Questions? Contact: divya@creativere.co")


# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center'>
    <p><small>Creative RE Underwriting Suite v1.0 | Built with Streamlit</small></p>
</div>
""", unsafe_allow_html=True)
