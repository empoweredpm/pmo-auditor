import streamlit as st
import pandas as pd
from openai import OpenAI

# 1. Professional Page Setup
st.set_page_config(page_title="Accidental PM Auditor", page_icon="🛡️", layout="wide")

# 2. Sidebar / Quickstart Menu
with st.sidebar:
    st.title("🛡️ Auditor Menu")
    st.markdown("""
    ### 🚀 Quick Start Guide
    1. **Export** your project plan to CSV or Excel.
    2. **Upload** the file using the box on the right.
    3. **Review** the AI-generated Risk Report.
    
    *Target Audience: Accidental Project Managers*
    """)
    st.divider()
    st.info("Version 1.0 - Production Ready")

# 3. Main Interface
st.title("🛡️ Accidental PM Auditor")
st.subheader("Turn your 'Accidental' project plan into a professional roadmap.")

# 4. API & Security Check
# This looks for the key you saved in the 'Advanced Settings' Secrets box
api_key = st.secrets.get("OPENAI_API_KEY", "")

if not api_key:
    st.error("🚨 **System Offline:** API Key not found in Streamlit Secrets.")
    st.stop() # Stops the app here so no code leaks
else:
    # 5. File Uploader (Now accepts Excel and CSV)
    uploaded_file = st.file_uploader("Upload Project Schedule", type=['csv', 'xlsx', 'xls'])

    if uploaded_file:
        try:
            # Handle both Excel and CSV
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ Loaded: {uploaded_file.name}")
            st.dataframe(df.head()) # Shows the first few rows of the plan
            
            st.button("🔍 Run Health Audit")
            
        except Exception as e:
            st.error(f"Error reading file: {e}")
