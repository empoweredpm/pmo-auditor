import streamlit as st
from openai import OpenAI

# 1. Basic Page Config
st.set_page_config(page_title="PMO Auditor", page_icon="🛡️")
st.title("🛡️ Accidental PM Auditor")

# 2. Secure API Key Logic
# This pulls the key from the hidden 'Secrets' vault we will set up next
api_key = st.secrets.get("OPENAI_API_KEY", "")

if not api_key:
    st.error("⚠️ System Offline: API Key missing from Secrets vault.")
    st.info("To fix: Go to Settings > Secrets in Streamlit and add: OPENAI_API_KEY = 'your-key'")
else:
    st.success("✅ System Online: Secure Connection Established.")
    client = OpenAI(api_key=api_key)

    # 3. Simple Interaction
    st.write("### Project Health Check")
    uploaded_file = st.file_uploader("Upload your project CSV", type="csv")
    
    if uploaded_file:
        st.info("File received. Ready to audit once we re-add the logic!")
