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
    """)
    st.divider()
    st.info("Target Audience: Accidental Project Managers")

# 3. Main Interface
st.title("🛡️ Accidental PM Auditor")
st.subheader("Automated Risk Analysis for Non-Certified Project Leaders")

# 4. API & Security Check
api_key = st.secrets.get("OPENAI_API_KEY", "")

if not api_key:
    st.error("🚨 **System Offline:** Please add your OPENAI_API_KEY to the Streamlit Secrets vault.")
    st.stop()
else:
    client = OpenAI(api_key=api_key)
    
    # 5. File Uploader
    uploaded_file = st.file_uploader("Upload Project Schedule (Excel or CSV)", type=['csv', 'xlsx', 'xls'])

    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ Loaded: {uploaded_file.name}")
            
            # 6. The "Original" Audit Logic
            if st.button("🔍 Run Full Health Audit"):
                with st.spinner("Analyzing project logic and risks..."):
                    # Convert the project plan to text for the AI
                    project_summary = df.to_string()
                    
                    response = client.chat.completions.create(
                        model="gpt-4",
                        messages=[
                            {"role": "system", "content": "You are a Senior PMO Auditor. Analyze this project plan for an 'Accidental PM'. Identify missing dependencies, unrealistic durations, and resource gaps."},
                            {"role": "user", "content": f"Review this project data: {project_summary}"}
                        ]
                    )
                    
                    st.markdown("### 📋 PMO Audit Report")
                    st.write(response.choices[0].message.content)
            
            st.divider()
            st.markdown("### Preview of Uploaded Plan")
            st.dataframe(df)
            
        except Exception as e:
            st.error(f"Error processing file: {e}")
