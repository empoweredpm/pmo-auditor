import streamlit as st
import pandas as pd
from openai import OpenAI
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os

# --- 1. PAGE SETUP ---
st.set_page_config(page_title="The Empowered PM Auditor", layout="wide", page_icon="🛡️")

# CSS: Styling for The Empowered PM branding
st.markdown("""
    <style>
    .stButton > button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; color: white !important; }
    [data-testid="stHorizontalBlock"] div:nth-child(1) button { background-color: #d32f2f !important; border: none; }
    [data-testid="stHorizontalBlock"] div:nth-child(2) button { background-color: #2e7d32 !important; border: none; }
    [data-testid="stHorizontalBlock"] div:nth-child(3) button { 
        background-color: #bdc3c7 !important; color: #000000 !important; border: 2px solid #2c3e50 !important;
    }
    .hero-section { background-color: #f0f4f8; padding: 25px; border-radius: 15px; border-left: 10px solid #0052cc; margin-bottom: 20px; }
    .sidebar-guide { font-size: 0.85rem; color: #333; background-color: #f0f2f6; padding: 12px; border-radius: 8px; border: 1px solid #d1d5db; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SECRETS & LOGIC ---
api_key = st.secrets.get("OPENAI_API_KEY", "")

def format_clean_text(doc, text):
    """Helper to remove asterisks and create real Word bullet points/spacing."""
    lines = text.split('\n')
    for line in lines:
        clean_line = line.replace('*', '').strip()
        if not clean_line:
            doc.add_paragraph() # Spacing
            continue
        
        # Check if the line was originally a bullet point
        if line.strip().startswith('*') or line.strip().startswith('-'):
            p = doc.add_paragraph(clean_line, style='List Bullet')
        else:
            p = doc.add_paragraph(clean_line)
        
        # Add spacing after paragraph
        p.paragraph_format.space_after = Pt(6)

# WORD DOC GENERATOR (Version 2.0: Clean Formatting & Watermark)
def create_word_doc(domain, audit, recovery):
    doc = Document()
    
    # 1. Header Logo
    if os.path.exists("empPMlogo.png"):
        doc.add_picture("empPMlogo.png", width=Inches(1.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. Branded Title
    title = doc.add_heading('The Empowered PM: Audit & Recovery Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    tagline = doc.add_paragraph('Effective PM Practices for Everyone')
    tagline.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tagline.runs[0].italic = True
    
    doc.add_heading('🕵️ Logic Audit Findings', level=1)
    format_clean_text(doc, audit)
    
    doc.add_heading('🛠️ Recovery Roadmap', level=1)
    format_clean_text(doc, recovery)
    
    # 3. Footer Watermark
    section = doc.sections[0]
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.text = "Prepared by The Empowered PM Consulting, copyright 2026."
    f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f_p.style.font.size = Pt(9)
    f_p.style.font.color.rgb = (128, 128, 128) # Grey watermark
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. MAIN INTERFACE HEADER ---
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if os.path.exists("empPMlogo.png"):
        st.image("empPMlogo.png", width=150)
with col_title:
    st.title("The Empowered PM Auditor")
    st.write("### *Effective PM Practices for Everyone*")

st.markdown("""
<div class="hero-section">
    <strong>The Logic Linter.</strong> Identify schedule errors and dependency gaps before they hit your critical path. 
    Focus on logic and clarity within your existing MS Office workflows.
</div>
""", unsafe_allow_html=True)

# --- 4. SIDEBAR ---
with st.sidebar:
    st.header("Auditor Control")
    project_context = st.selectbox("Project Domain", ["IT/Software", "Construction", "Operations", "Marketing", "Events"])
    
    st.divider()
    st.markdown("**📖 Quick Start Guide**")
    st.markdown("""
    <div class="sidebar-guide">
    1. <b>Upload:</b> Plan in XLSX or CSV format.<br>
    2. <b>Audit (Red):</b> Execute logic critique.<br>
    3. <b>Recover (Green):</b> Generate fixes.<br>
    4. <b>Export:</b> Download the clean Word report.
    </div>
    """, unsafe_allow_html=True)

# --- 5. MAIN INTERFACE LOGIC ---
if not api_key:
    st.error("🚨 API Key missing from Secrets.")
    st.stop()

uploaded_file = st.file_uploader("Upload Project Schedule (XLSX or CSV)", type=["xlsx", "csv"])

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
                        messages=[
                            {"role": "system", "content": f"You are a Senior PMO Auditor for {project_context}. Audit this schedule for logic errors. IMPORTANT: Do NOT recommend new PM tools (Jira, Asana, etc.). Assume they use Excel/MS Project. Provide a health score (0-100%) and plain text math only (no LaTeX)."},
                            {"role": "user", "content": schedule_data}
                        ]
                    )
                    st.session_state['audit_report'] = response.choices[0].message.content
                    st.rerun()
        
        with col2:
            if 'audit_report' in st.session_state:
                if st.button("🛠️ GENERATE RECOVERY PLAN"):
                    with st.spinner("Building roadmap..."):
                        response = client.chat.completions.create(
                            model="gpt-4-turbo",
                            messages=[
                                {"role": "system", "content": "Create a 3-step recovery roadmap based on findings. Focus on fixing dates and logic in MS Office applications. DO NOT suggest software migrations. Use clean text without markdown asterisks."},
                                {"role": "user", "content": st.session_state['audit_report']}
                            ]
                        )
                        st.session_state['recovery_plan'] = response.choices[0].message.content
                        st.rerun()
        
        with col3:
            if st.button("🗑️ RESET ALL DATA"):
                for key in ['audit_report', 'recovery_plan']:
                    if key in st.session_state: del st.session_state[key]
                st.rerun()

        if 'audit_report' in st.session_state:
            st.divider()
            st.subheader("🕵️ Auditor's Findings")
            st.markdown(st.session_state['audit_report'])
            
        if 'recovery_plan' in st.session_state:
            st.success("### ✅ Practical Recovery Roadmap")
            st.markdown(st.session_state['recovery_plan'])
            
            word_data = create_word_doc(project_context, st.session_state['audit_report'], st.session_state['recovery_plan'])
            st.download_button(label="📥 Download Report", data=word_data, file_name="Empowered_PM_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    except Exception as e:
        st.error(f"Error: {e}")
