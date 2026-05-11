# THS Templates

## Setup Instructions

1. **Add your THS template here:**
   - Place your pre-loaded THS template file in this folder
   - Name it: `THS_Template_Default.xlsx`
   - This template should have sheets: `Trend Health Score`, `T12`, `Rent Roll`, `Anchors & Notes`

2. **The app will automatically load:** `THS_Template_Default.xlsx` at startup

3. **To update the template:**
   - Replace the file in this folder
   - No code changes needed
   - App will use the new template on next run

## Expected Structure

Your THS template should have:
- **Sheet 1: Trend Health Score** - Main scoring sheet (C2=Property, C4=Address, C6=Date)
- **Sheet 2: T12** - Where T12 data gets populated
- **Sheet 3: Rent Roll** - Where RR metrics go
- **Sheet 4: Anchors & Notes** - Scoring guidance

The app will populate all data into these sheets and download a single file.
