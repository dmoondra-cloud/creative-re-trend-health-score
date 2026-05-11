# 🚀 Quick Setup Guide

## Before First Use: Add Your THS Template

The app requires your **pre-loaded THS template** to be present.

### Step 1: Get Your Template Ready
- Use the THS template you provided earlier
- It should have these sheets:
  - `Trend Health Score` (main sheet)
  - `T12` (for line items)
  - `Rent Roll` (for unit data)
  - `Anchors & Notes` (scoring guide)

### Step 2: Copy to Templates Folder
1. Navigate to your app folder: `uw_app/`
2. Inside, you'll see a `templates/` folder
3. **Copy your THS template file into `templates/` folder**
4. **Rename it to:** `THS_Template_Default.xlsx`

```
uw_app/
├── app.py
├── templates/
│   └── THS_Template_Default.xlsx  ← Put your template here
├── categorization_engine.py
└── ... (other files)
```

### Step 3: Run the App

**Locally:**
```bash
cd uw_app
pip install -r requirements.txt
streamlit run app.py
```

Then open: `http://localhost:8501`

**On Streamlit Cloud:**
1. Push everything to GitHub (including the template in templates/ folder)
2. Go to share.streamlit.io
3. Create new app → Select repo & `app.py`
4. It will auto-deploy

---

## App Workflow

1. **Enter Property Details**
   - Property name, address, URL

2. **Upload Files**
   - Raw T12 (any format: Yardi, MRI, etc.)
   - Raw Rent Roll
   - RR as-of date
   - Floor plans (optional)

3. **Review & Edit**
   - Left: T12 categorization (with edit dropdowns)
   - Right: RR summary (preview of data)
   - Shows if any amounts were multiplied by ±1

4. **Approve & Generate**
   - Click "Approve & Generate THS"
   - App populates everything into your THS template
   - **Single download:** Your completed THS workbook

---

## What Gets Populated

Your **downloaded THS file** will have:

| Sheet | Content |
|-------|---------|
| Trend Health Score | Property info + metrics |
| T12 | All categorized line items + values |
| Rent Roll | Unit data + metrics |
| Anchors & Notes | Scoring guidance (unchanged) |

---

## Troubleshooting

**Error: "Template not found"**
- Make sure file is at: `uw_app/templates/THS_Template_Default.xlsx`
- Check filename is exact (case-sensitive on some systems)

**Error: "Could not parse T12"**
- Make sure T12 file is .xlsx or .xls
- File should have clear month headers and line items

**RR not showing data**
- Check RR file format is correct
- Make sure it has a sheet named "Rent Roll" or "RR"

---

## Questions?

Contact: divya@creativere.co

---

Built with ❤️ for Creative RE Underwriting
