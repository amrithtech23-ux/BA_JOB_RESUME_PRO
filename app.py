import streamlit as st
import requests
import os
from pypdf import PdfReader
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT
from datetime import datetime
import tempfile

# Page Configuration
st.set_page_config(
    page_title="Business Analyst Job Apply Pro",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .report-box {
        font-family: 'Courier New', Courier, monospace;
        white-space: pre;
        background-color: #0e1117;
        color: #00ff00;
        padding: 20px;
        border-radius: 10px;
        overflow-x: auto;
    }
    </style>
""", unsafe_allow_html=True)

# --- Helper Functions ---

def extract_text_from_file(uploaded_file):
    """Extracts text from PDF, DOCX, or TXT files."""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PdfReader(uploaded_file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""
            return text.strip()
        elif uploaded_file.name.endswith('.docx'):
            doc = Document(uploaded_file)
            return "\n".join([para.text for para in doc.paragraphs])
        elif uploaded_file.name.endswith('.txt'):
            return uploaded_file.read().decode("utf-8")
        else:
            return None
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return None

def get_api_key():
    """Get API key from secrets or user input."""
    # Try to get from secrets first
    try:
        if hasattr(st, 'secrets') and "OPENROUTER_API_KEY" in st.secrets:
            return st.secrets["OPENROUTER_API_KEY"]
    except:
        pass
    
    # Fallback to user input
    return st.text_input("OpenRouter API Key", type="password", 
                        help="Get your key from https://openrouter.ai/keys")

def call_openrouter_api(prompt, system_instruction, api_key):
    """Calls OpenRouter API using Qwen model."""
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/amrithtech23-ux/BA_JOB_RESUME_PRO",
        "X-Title": "BA Job Apply Pro"
    }
    
    payload = {
        "model": "qwen/qwen-2.5-72b-instruct",
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3,
        "max_tokens": 4000
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        st.error(f"API Error: {str(e)}")
        return None

def generate_pdf_resume(resume_text):
    """Generates a simple ATS-friendly PDF."""
    try:
        # Create a temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_filename = temp_file.name
        temp_file.close()
        
        doc = SimpleDocTemplate(temp_filename, pagesize=A4,
                               rightMargin=0.75*72, leftMargin=0.75*72,
                               topMargin=0.75*72, bottomMargin=0.75*72)
        styles = getSampleStyleSheet()
        
        style_normal = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=12,
            alignment=TA_LEFT,
            spaceAfter=6
        )
        
        style_heading = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=14,
            leading=16,
            spaceBefore=12,
            spaceAfter=6
        )
        
        story = []
        lines = resume_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 6))
                continue
            
            if line.startswith("###"):
                story.append(Paragraph(line.replace("###", "").strip(), style_heading))
            else:
                story.append(Paragraph(line, style_normal))
        
        doc.build(story)
        return temp_filename
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

# --- Prompts ---

VALIDATION_SYSTEM_PROMPT = """You are an expert IT Business Analyst Hiring Manager. 
Compare the Candidate's Resume against the Job Description.
Output the result in the exact ASCII ART table format shown in the example.
Be specific about matches and gaps."""

RESUME_GEN_SYSTEM_PROMPT = """You are an expert Resume Writer specializing in ATS-optimized IT Business Analyst resumes.
Rewrite the user's resume to align with the Job Description.
Follow the exact structure of the Priya Sharma template:
1. Name
2. Title
3. Contact Info
4. Professional Summary
5. Domain Expertise
6. Professional Experience
7. Certifications
8. Technical & Professional Skills
9. Key Projects
10. Education
Use single column, no graphics, no tables."""

# --- Main App ---

def main():
    st.title("📊 Business Analyst Job Apply Pro")
    st.markdown("### ATS Optimized Resume Validator & Generator")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Get API key
        api_key = get_api_key()
        
        if not api_key:
            st.warning("⚠️ Please enter your OpenRouter API Key")
            st.stop()
        else:
            st.success("✅ API Key configured")
        
        st.markdown("---")
        st.markdown("### About")
        st.markdown("Powered by OpenRouter API & Qwen Models")
    
    # Main content
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1️⃣ Upload Resume")
        uploaded_file = st.file_uploader("Choose file", type=['pdf', 'docx', 'txt'])
        
        st.subheader("2️⃣ Job Description")
        jd_text = st.text_area("Paste JD here", height=200)
    
    with col2:
        st.subheader("3️⃣ Actions")
        validate_btn = st.button("🔍 Validate Resume", type="primary")
        reset_btn = st.button("🔄 Reset")
        
        if reset_btn:
            st.session_state.clear()
            st.rerun()
    
    # Session state
    if 'validation_report' not in st.session_state:
        st.session_state.validation_report = None
    if 'generated_resume' not in st.session_state:
        st.session_state.generated_resume = None
    
    # Process
    if uploaded_file and jd_text:
        # Extract resume
        resume_text = extract_text_from_file(uploaded_file)
        
        if not resume_text:
            st.error("❌ Could not extract text from resume")
            st.stop()
        
        if validate_btn:
            with st.spinner("Analyzing..."):
                prompt = f"""
Resume:
{resume_text}

Job Description:
{jd_text}

Generate validation report in ASCII table format."""
                
                report = call_openrouter_api(prompt, VALIDATION_SYSTEM_PROMPT, api_key)
                
                if report:
                    st.session_state.validation_report = report
                    st.success("✅ Validation complete!")
        
        # Show report
        if st.session_state.validation_report:
            st.markdown("### 📋 Validation Report")
            st.code(st.session_state.validation_report, language="text")
            
            st.markdown("---")
            if st.button("✨ Generate Optimized Resume"):
                with st.spinner("Generating..."):
                    prompt = f"""
Original Resume:
{resume_text}

Job Description:
{jd_text}

Generate ATS-optimized resume following Priya Sharma template structure."""
                    
                    new_resume = call_openrouter_api(prompt, RESUME_GEN_SYSTEM_PROMPT, api_key)
                    
                    if new_resume:
                        st.session_state.generated_resume = new_resume
                        st.success("✅ Resume generated!")
        
        # Show generated resume
        if st.session_state.generated_resume:
            st.markdown("### 📄 Generated Resume")
            st.text_area("Preview", st.session_state.generated_resume, height=400)
            
            # Generate PDF
            pdf_file = generate_pdf_resume(st.session_state.generated_resume)
            
            if pdf_file:
                with open(pdf_file, "rb") as f:
                    st.download_button(
                        label="📥 Download PDF",
                        data=f,
                        file_name="ATS_Resume.pdf",
                        mime="application/pdf"
                    )
                
                # Cleanup
                try:
                    os.unlink(pdf_file)
                except:
                    pass
    
    elif uploaded_file and not jd_text:
        st.info("📝 Please paste the job description")
    elif jd_text and not uploaded_file:
        st.info("📄 Please upload your resume")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"❌ Application error: {str(e)}")
        st.exception(e)  # This will show the full traceback
