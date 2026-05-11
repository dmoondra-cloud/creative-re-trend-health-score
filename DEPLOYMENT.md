# Deployment Guide

## 🚀 Deploy to Streamlit Cloud (Recommended)

### Prerequisites
- GitHub account
- Streamlit Community Cloud account (free)
- Repository access

### Step 1: Push to GitHub

```bash
# Initialize git repo (if not already done)
git init
git add .
git commit -m "Initial commit: Creative RE UW Suite"

# Add remote repository
git remote add origin https://github.com/YOUR_USERNAME/uw_app.git

# Push to GitHub
git branch -M main
git push -u origin main
```

### Step 2: Connect to Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click "New app"
3. Select your GitHub repository
4. Select branch: `main`
5. Set main file path: `app.py`
6. Click "Deploy"

### Step 3: Configure Secrets (Optional)

If using API keys or sensitive data:

1. In Streamlit Cloud dashboard, go to app settings
2. Click "Secrets"
3. Add secrets in TOML format:

```toml
[database]
connection_string = "your_connection_string"

[api]
key = "your_api_key"
```

Access in app:
```python
import streamlit as st
secret = st.secrets["database"]["connection_string"]
```

---

## 💻 Deploy Locally

### Option 1: Direct Installation

```bash
# Install Python 3.8+
python --version

# Clone repository
git clone https://github.com/YOUR_USERNAME/uw_app.git
cd uw_app

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run app
streamlit run app.py
```

App will be available at: `http://localhost:8501`

### Option 2: Docker

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

CMD ["streamlit", "run", "app.py"]
```

Build and run:
```bash
docker build -t uw_app .
docker run -p 8501:8501 uw_app
```

---

## 🔧 Development Environment

### Setup Development

```bash
# Clone for development
git clone https://github.com/YOUR_USERNAME/uw_app.git
cd uw_app

# Create venv
python -m venv venv
source venv/bin/activate

# Install dependencies + dev tools
pip install -r requirements.txt
pip install pytest black flake8

# Run tests
pytest tests/

# Format code
black *.py

# Lint
flake8 *.py
```

### Testing

Create `tests/test_parser.py`:
```python
import pytest
from t12_parser import T12Parser

def test_parse_basic():
    parser = T12Parser('sample_t12.xlsx')
    result = parser.parse()
    assert result['parsed_successfully']
    assert result['property_name'] != 'Unknown Property'
```

Run tests:
```bash
pytest tests/
```

---

## 📦 Updating the App

### Make Changes Locally

1. Edit files locally
2. Test thoroughly:
   ```bash
   streamlit run app.py
   ```
3. Commit and push:
   ```bash
   git add .
   git commit -m "Describe changes"
   git push origin main
   ```

4. **Streamlit Cloud auto-redeploys** when you push to `main`

### Version Management

Use Git tags for releases:
```bash
git tag -a v1.0.0 -m "First release"
git push origin v1.0.0
```

---

## 🔐 Security Checklist

- [ ] Don't commit API keys or secrets
- [ ] Use `.gitignore` to exclude sensitive files
- [ ] Use Streamlit Cloud Secrets for credentials
- [ ] Validate all user inputs
- [ ] Don't write temp files to cloud storage
- [ ] Use HTTPS for any API calls
- [ ] Regular dependency updates: `pip list --outdated`

---

## 🐛 Troubleshooting

### App Won't Deploy
- Check `requirements.txt` - all imports must be listed
- Verify Python version compatibility (3.8+)
- Check file paths (use relative paths)
- Review Streamlit Cloud build logs

### Memory Issues
- Don't load entire files into memory
- Use streaming for large datasets
- Clear session state if needed: `st.session_state.clear()`

### Slow Performance
- Optimize T12 parsing logic
- Cache expensive operations: `@st.cache_data`
- Use `st.spinner()` for long tasks

### Template Issues
- Verify Excel file format (not .xls)
- Check sheet names match `TEMPLATE_SHEETS`
- Validate cell references in formulas

---

## 📊 Monitoring

### Streamlit Cloud Analytics
- Dashboard shows app usage
- View recent deployments
- Check error logs in "Manage App"

### Local Logging
Add logging for debugging:
```python
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

logger.info("T12 parsing started")
```

---

## 🚀 Production Checklist

Before going live:
- [ ] All tests passing
- [ ] Documentation complete
- [ ] Error handling robust
- [ ] Security validated
- [ ] Performance optimized
- [ ] Backup/recovery plan
- [ ] User training materials
- [ ] Support contact info

---

For questions: **divya@creativere.co**
