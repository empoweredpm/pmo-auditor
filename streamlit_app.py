import streamlit as st
import pandas as pd
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- 1. PAGE SETUP & BRANDING ---
st.set_page_config(page_title="The Empowered PM Auditor", layout="wide", page_icon="🛡️")

st.markdown("""
    <style>
    .stButton > button { width: 100%; border-radius: 8px; height: 3.5em; font-weight: bold; color: white !important; }
    [data-testid="stHorizontalBlock"] div:nth-child(1) button { background-color: #d32f2f !important; }
    [data-testid="stHorizontalBlock"] div:nth-child(2) button { background-color: #2e7d32 !important; }
    [data-testid="stHorizontalBlock"] div:nth-child(3) button { background-color: #bdc3c7 !important; color: black !important; border: 2px solid #2c3e50 !important; }
    
    .hero-section { background-color: #f0f4f8; padding: 30px; border-radius: 15px; border-left: 10px solid #0052cc; margin-bottom: 25px; }
    .tagline { color: #0052cc; font-weight: bold; font-size: 1.2rem; margin-top: -15px; margin-bottom: 20px; }
    .coaching-box { background-color: #fff3cd; padding: 20px; border-radius: 10px; border: 1px solid #ffeeba; margin-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. BACKGROUND API LOGIC ---
api_key = st.secrets.get("OPENAI_API_KEY", "")

# BRANDED WORD DOC GENERATOR
def create_word_doc(domain, audit, recovery):
    doc = Document()
    
    # Branded Header
    header = doc.add_heading('The Empowered PM: Audit Report', 0)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Effective PM Practices for Everyone')
    run.italic = True
    run.font.size = Pt(11)
    
    doc.add_heading('🕵️ Logic Audit Findings', level=1)
    doc.add_paragraph(audit)
    
    doc.add_heading('🛠️ Recovery Roadmap', level=1)
    doc.add_paragraph(recovery)
    
    # Professional Call to Action
    doc.add_page_break()
    cta = doc.add_heading('Next Steps: Join the 4-Week PM Survival Intensive', level=1)
    doc.add_paragraph("Stop just managing tasks. Start leading projects with authority. This audit is your first step toward becoming an Empowered PM.")
    doc.add_paragraph("Visit: [Your Website Link]")
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. SIDEBAR (The Empowered PM Control) ---
with st.sidebar:
    # --- LOGO PLACEHOLDER ---
    # To use your logo, upload it to GitHub as 'logo.png' then uncomment the line below:
    # st.image("logo.png", use_column_width=True)
    
    st.header("🛡️ Auditor Control")
    project_context = st.selectbox("Project Domain", ["IT/Software", "Construction", "Operations", "Marketing", "Event Management"])
    
    st.markdown("""
    <div class="coaching-box">
    <strong>🚀 The Empowered PM Intensive</strong><br>
    Ready to lead with influence? Join my 4-week survival program for Accidental PMs.
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    st.info("Goal: Transition from Accidental PM to Empowered Leader.")

# --- 4. MAIN INTERFACE (The Landing Page) ---
st.markdown("""
<div class="hero-section">
    <h1>The Empowered PM Auditor</h1>
    <div class="tagline">Effective PM Practices for Everyone</div>
    <p>Our <strong>Logic Linter</strong> identifies schedule bugs and risk gaps 
    so you can present your project plan with absolute confidence.</p>
</div>
""", unsafe_allow_html=True)

if not api_key:
    st.error("API Key missing.")
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
                with st.spinner("Analyzing plan logic..."):
                    schedule_data = df.to_string(index=False)
                    response = client.chat.completions.create(
                        model="gpt-4-turbo",
                        messages=[
                            {"role": "system", "content": "You are The Empowered PM Mentor. Audit this schedule for logic errors. Be authoritative yet supportive. Score 0-100%."},
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
                                {"role": "system", "content": "Create a 3-step recovery roadmap based on these findings. Focus on 'Empowered PM' principles: Clarity, Influence, and Communication."},
                                {"role": "user", "content": st.session_state['audit_report']}
                            ]
                        )
                        st.session_state['recovery_plan'] = response.choices[0].message.content
                        st.rerun()
        
        with col3:
            if st.button("🗑️ RESET ALL"):
                for key in list(st.session_state.keys()): del st.session_state[key]
                st.rerun()

        if 'audit_report' in st.session_state:
            st.divider()
            st.subheader("🕵️ Logic Audit Findings")
            st.markdown(st.session_state['audit_report'])
            
        if 'recovery_plan' in st.session_state:
            st.success("### ✅ Your Empowered Recovery Roadmap")
            st.markdown(st.session_state['recovery_plan'])
            
            word_data = create_word_doc(project_context, st.session_state['audit_report'], st.session_state['recovery_plan'])
            st.download_button("📥 Download Branded Recovery Report (.docx)", data=word_data, file_name="Empowered_PM_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            
    except Exception as e:
        st.error(f"Error: {e}")
