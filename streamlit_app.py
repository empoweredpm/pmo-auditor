import streamlit as st
import pandas as pd
from openai import OpenAI
import re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="The Empowered PM Auditor", layout="wide", page_icon="🛡️")

# Custom CSS for the Agency Banner and Buttons
st.markdown("""
    <style>
    .stButton > button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; color: white !important; }
    
    /* Audit (Red), Recovery (Green), Reset (Grey) */
    [data-testid="stHorizontalBlock"] div:nth-child(1) button { background-color: #d32f2f !important; border: none; }
    [data-testid="stHorizontalBlock"] div:nth-child(2) button { background-color: #2e7d32 !important; border: none; }
    [data-testid="stHorizontalBlock"] div:nth-child(3) button { 
        background-color: #bdc3c7 !important; color: #000000 !important; border: 2px solid #2c3e50 !important;
    }

    .hero-section { background-color: #f0f4f8; padding: 25px; border-radius: 15px; border-left: 10px solid #0052cc; margin-bottom: 20px; }
    
    /* Agency Banner / CTA Box */
    .agency-banner { 
        background-color: #1a2a6c; 
        color: white; 
        padding: 30px; 
        border-radius: 12px; 
        text-align: center; 
        margin-top: 40px;
        border: 2px solid #fdbb2d;
    }
    .agency-banner h2 { color: #fdbb2d !important; margin-bottom: 10px; }
    .cta-button {
        display: inline-block;
        padding: 12px 24px;
        margin: 10px;
        border-radius: 5px;
        text-decoration: none;
        font-weight: bold;
        font-size: 16px;
    }
    .btn-rescue { background-color: #fdbb2d; color: #1a2a6c !important; }
    .btn-course { background-color: white; color: #1a2a6c !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SECRETS & LOGIC ---
api_key = st.secrets.get("OPENAI_API_KEY", "")

if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

def format_clean_text(doc, text):
    lines = text.split('\n')
    for line in lines:
        clean_line = line.replace('*', '').strip()
        if not clean_line:
            doc.add_paragraph()
            continue
        if line.strip().startswith('*') or line.strip().startswith('-'):
            doc.add_paragraph(clean_line, style='List Bullet')
        else:
            doc.add_paragraph(clean_line)

def create_word_doc(domain, audit, recovery):
    doc = Document()
    # Add logo to doc if exists
    if os.path.exists("empPMlogo.jpg"):
        doc.add_picture("empPMlogo.jpg", width=Inches(1.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    title = doc.add_heading('The Empowered PM: Audit & Recovery Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('🕵️ Audit Findings', level=1)
    format_clean_text(doc, audit)
    
    doc.add_heading('🛠️ Recovery Roadmap', level=1)
    format_clean_text(doc, recovery)
    
    section = doc.sections[0]
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.text = "Prepared by The Empowered PM Consulting | theempoweredpm.com"
    f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. INTERFACE ---
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if os.path.exists("empPMlogo.jpg"):
        st.image("empPMlogo.jpg", width=150)
with col_title:
    st.title("The Empowered PM Auditor")
    st.write("### *Effective PM Practices for Everyone*")

st.markdown("""
<div class="hero-section">
    <strong>Validation & Practical Recovery.</strong> Identify schedule errors and dependency gaps 
    to lead your project with absolute clarity and authority.
</div>
""", unsafe_allow_html=True)

# --- 4. SIDEBAR ---
with st.sidebar:
    st.header("Auditor Control")
    project_context = st.selectbox("Project Domain", ["IT/Software", "Construction", "Operations", "Marketing", "Events"])
    st.divider()
    st.markdown("**📖 Quick Start Guide**")
    st.info("1. Upload Schedule\n2. Run Audit\n3. Generate Recovery\n4. Export Report")

# --- 5. MAIN LOGIC ---
if not api_key:
    st.error("🚨 API Key missing from Secrets.")
    st.stop()

uploaded_file = st.file_uploader("Upload Project Schedule (XLSX or CSV)", type=["xlsx", "csv"], key=f"uploader_{st.session_state['uploader_key']}")

if uploaded_file:
    client = OpenAI(api_key=api_key)
    try:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        st.subheader("📋 Active Schedule Data")
        st.dataframe(df, use_container_width=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("🚀 EXECUTE LOGIC AUDIT"):
                with st.spinner("Auditing logic..."):
                    schedule_data = df.to_string(index=False)
                    response = client.chat.completions.create(
                        model="gpt-4-turbo",
                        messages=[{"role": "system", "content": "You are a Senior PMO Auditor. Audit for logic errors. Plain text only."}, {"role": "user", "content": schedule_data}]
                    )
                    st.session_state['audit_report'] = response.choices[0].message.content
                    st.rerun()
        
        with col2:
            if 'audit_report' in st.session_state:
                if st.button("🛠️ GENERATE RECOVERY PLAN"):
                    with st.spinner("Building roadmap..."):
                        response = client.chat.completions.create(
                            model="gpt-4-turbo",
                            messages=[{"role": "system", "content": "Create a 3-step recovery roadmap based on the audit. No asterisks."}, {"role": "user", "content": st.session_state['audit_report']}]
                        )
                        st.session_state['recovery_plan'] = response.choices[0].message.content
                        st.rerun()
        
        with col3:
            if st.button("🗑️ RESET ALL DATA"):
                if 'audit_report' in st.session_state: del st.session_state['audit_report']
                if 'recovery_plan' in st.session_state: del st.session_state['recovery_plan']
                st.session_state['uploader_key'] += 1
                st.rerun()

        # Display Results
        if 'audit_report' in st.session_state:
            st.divider()
            st.subheader("🕵️ Auditor's Findings")
            st.markdown(st.session_state['audit_report'])
            
        if 'recovery_plan' in st.session_state:
            st.success("### ✅ Your Empowered Recovery Roadmap")
            st.markdown(st.session_state['recovery_plan'])
            
            # --- THE AGENCY AGENT BANNER (CTA) ---
            st.markdown(f"""
                <div class="agency-banner">
                    <h2>🛡️ Need a Professional Rescue?</h2>
                    <p>Don't present a broken plan. Let's fix your logic together or join our training cohort.</p>
                    <a href="https://calendly.com/empoweredpming/rescue_session" class="cta-button btn-rescue" target="_blank">Book 15-Min Rescue Call</a>
                    <a href="https://empoweredpm.com" class="cta-button btn-course" target="_blank">Join 4-Week Intensive</a>
                </div>
            """, unsafe_allow_html=True)
            
            # Download Button
            word_data = create_word_doc(project_context, st.session_state['audit_report'], st.session_state['recovery_plan'])
            st.download_button(label="📥 Download Full Report", data=word_data, file_name="Empowered_PM_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            
    except Exception as e:
        st.error(f"Error processing file: {e}")

st.divider()
st.caption("© 2026 The Empowered PM Consulting")
