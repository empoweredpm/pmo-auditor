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

# Initialize File Uploader Key if not present
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

def format_clean_text(doc, text):
    """Helper to remove asterisks and create real Word bullet points/spacing."""
    lines = text.split('\n')
    for line in lines:
        clean_line = line.replace('*', '').strip()
        if not clean_line:
            doc.add_paragraph()
            continue
        
        if line.strip().startswith('*') or line.strip().startswith('-'):
            p = doc.add_paragraph(clean_line, style='List Bullet')
        else:
            p = doc.add_paragraph(clean_line)
        p.paragraph_format.space_after = Pt(6)

def create_word_doc(domain, audit, recovery):
    doc = Document()
    if os.path.exists("empPMlogo.png"):
        doc.add_picture("empPMlogo.png", width=Inches(1.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    title = doc.add_heading('The Empowered PM: Audit & Recovery Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    tagline = doc.add_paragraph('Effective PM Practices for Everyone')
    tagline.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tagline.runs[0].italic = True
    
    doc.add_heading('🕵️ Logic Audit Findings', level=1)
    format_clean_text(doc, audit)
    
    doc.add_heading('🛠️ Recovery Roadmap', level=1)
    format_clean_text(doc, recovery)
    
    # Watermark Footer
    section = doc.sections[0]
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.text = "Prepared by The Empowered PM Consulting, copyright 2026."
    f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f_p.style.font.size = Pt(9)
    f_p.style.font.color.rgb = (128, 128, 128)
    
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
    <strong>The Logic Linter.</strong> Clear your project data at any time using the Reset button. 
    Audit schedule errors and dependency gaps within your existing MS Office workflows.
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
    1. <b>Upload:</b> Schedule in XLSX or CSV.<br>
    2. <b>Audit (Red):</b> Execute logic critique.<br>
    3. <b>Recover (Green):</b> Generate fixes.<br>
    4. <b>Export:</b> Download clean report.<br>
    5. <b>Reset:</b> Wipe all data for next project.
    </div>
    """, unsafe_allow_html=True)

# --- 5. MAIN INTERFACE LOGIC ---
if not api_key:
    st.error("🚨 API Key missing from Secrets.")
    st.stop()

# File Uploader with Dynamic Key for Reset functionality
uploaded_file = st.file_uploader(
    "Upload Project Schedule (XLSX or CSV)", 
    type=["xlsx", "csv"], 
    key=f"uploader_{st.session_state['uploader_key']}"
)

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
                            {"role": "system", "content": f"You are a Senior PMO Auditor for {project_context}. Audit this schedule for logic errors. Do NOT recommend new PM tools. Assume MS Office use. Use plain text math only."},
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
                                {"role": "system", "content": "Create a 3-step recovery roadmap. Focus on MS Office fixes. No software migrations. No asterisks."},
                                {"role": "user", "content": st.session_state['audit_report']}
                            ]
                        )
                        st.session_state['recovery_plan'] = response.choices[0].message.content
                        st.rerun()
        
        with col3:
            if st.button("🗑️ RESET ALL DATA"):
                # 1. Clear text reports
                if 'audit_report' in st.session_state: del st.session_state['audit_report']
                if 'recovery_plan' in st.session_state: del st.session_state['recovery_plan']
                
                # 2. Increment uploader key to force clear the file widget
                st.session_state['uploader_key'] += 1
                
                # 3. Refresh the page
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

# --- 6. GLOSSARY ---
st.divider()
with st.expander("📚 The Empowered PM's Glossary"):
    st.markdown("""<div class="glossary-card"><strong>Logic Linter:</strong> A tool that scans your project plan for 'logic bugs' like missing links or impossible dates.</div>""", unsafe_allow_html=True)
