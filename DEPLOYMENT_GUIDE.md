# DTC to ClickUp Converter - Deployment Guide

## Quick Deploy to Streamlit Community Cloud (Free)

### Prerequisites
- GitHub account
- Your existing `dtc_to_clickup.py` script

### Step-by-Step Deployment (15 minutes)

#### 1. Create GitHub Repository
```bash
# On your local machine, create a new folder
mkdir dtc-clickup-converter
cd dtc-clickup-converter

# Copy your files
# - dtc_to_clickup.py (your existing script)
# - dtc_streamlit_app.py (the new Streamlit app)
# - requirements.txt

# Initialize git and push to GitHub
git init
git add .
git commit -m "Initial commit - DTC to ClickUp converter"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/dtc-clickup-converter.git
git push -u origin main
```

#### 2. Deploy to Streamlit Cloud
1. Go to https://share.streamlit.io
2. Sign in with GitHub
3. Click "New app"
4. Select your repository: `dtc-clickup-converter`
5. Set main file: `dtc_streamlit_app.py`
6. Click "Deploy!"

**That's it!** Your app will be live at: `https://YOUR_USERNAME-dtc-clickup-converter.streamlit.app`

---

## File Structure

Your repository should contain exactly 3 files:

```
dtc-clickup-converter/
├── dtc_to_clickup.py          # Your existing conversion logic (unchanged)
├── dtc_streamlit_app.py        # New Streamlit web interface
└── requirements.txt            # Python dependencies
```

---

## Local Testing (Optional)

Before deploying, test locally:

```bash
# Install dependencies
pip install -r requirements.txt

# Run locally
streamlit run dtc_streamlit_app.py
```

App opens at: http://localhost:8501

---

## Sharing with Team

Once deployed, share the URL with your team:
- No Python installation needed
- No command line needed
- Works on any device with a browser
- Free forever for internal use

**Example URL pattern:**
`https://arcadia-dtc-clickup.streamlit.app`

---

## Troubleshooting

**Issue: "Module not found" error**
- Check that `dtc_to_clickup.py` is in the same directory as `dtc_streamlit_app.py`

**Issue: "Conversion failed"**
- Verify Excel file contains sheets named like "Wk6", "Wk7", etc.
- Check that sheets follow expected format (see ANALYSIS.txt)

**Issue: Streamlit Cloud deployment fails**
- Verify all 3 files are committed to GitHub
- Check that requirements.txt exists
- Ensure repository is public or you have Streamlit Cloud private repo access

---

## Cost

- **Streamlit Community Cloud**: $0/month (free tier)
- **Compute**: Included
- **Storage**: Temporary only (files deleted after conversion)
- **Users**: Unlimited team access

---

## Maintenance

The app auto-deploys when you push changes to GitHub:

```bash
# Make changes to dtc_to_clickup.py or dtc_streamlit_app.py
git add .
git commit -m "Update conversion logic"
git push

# Streamlit Cloud auto-redeploys in ~2 minutes
```

---

## Alternative: Deploy to Netlify (Not Recommended for This Use Case)

Netlify requires additional configuration for Python apps:
- Need to containerize the app
- More complex setup than Streamlit Cloud
- No native Python support (requires serverless functions)

**Verdict**: Streamlit Cloud is the right choice for this tool.

---

## Questions?

Contact: Pete.buzzell@arcadiadigital.com
