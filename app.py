import streamlit as st
import requests
import json
import re
from pypdf import PdfReader
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT
import os

# Page Configuration
st.set_page_config(
    page_title="Business Analyst Job Apply Pro",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better ASCII rendering
st.markdown("""
    <style>
    .stCodeBlock {
        background-color: #f0f2f6;
        border: 1px solid #ddd;
        border-radius: 5px;
    }
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
                text += page.extract_text() + "\n"
            return text
        elif uploaded_file.name.endswith('.docx'):
            doc = Document(uploaded_file)
            return "\n".join([para.text for para in doc.paragraphs])
        elif uploaded_file.name.endswith('.txt'):
            return uploaded_file.read().decode("utf-8")
        else:
            st.error("Unsupported file format. Please upload PDF, DOCX, or TXT.")
            return None
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def call_openrouter_api(prompt, system_instruction, api_key):
    """Calls OpenRouter API using Qwen model."""
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/yourusername/ba-job-apply-pro", 
        "X-Title": "BA Job Apply Pro"
    }
    
    # Using Qwen model as per request context
    payload = {
        "model": "qwen/qwen-2.5-72b-instruct", 
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content']
    except Exception as e:
        st.error(f"API Error: {e}")
        return None

def generate_pdf_resume(resume_text, filename="Updated_Resume.pdf"):
    """Generates a simple ATS-friendly PDF based on the text content."""
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=0.75*inch, leftMargin=0.75*inch, topMargin=0.75*inch, bottomMargin=0.75*inch)
    styles = getSampleStyleSheet()
    
    # Custom Styles to match ATS Template (Single column, standard fonts)
    style_name = ParagraphStyle(
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
        spaceAfter=6,
        textColor='#000000'
    )
    
    style_subheading = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading3'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=13,
        spaceBefore=6,
        spaceAfter=2,
        textColor='#000000'
    )

    story = []
    
    # Simple parsing logic to format the generated text into PDF elements
    # This assumes the LLM returns structured text with headers marked by ### or **
    lines = resume_text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6))
            continue
        
        # Detect Headers based on common markdown-like patterns from LLM
        if line.startswith("###"):
            story.append(Paragraph(line.replace("###", "").strip(), style_heading))
        elif line.startswith("**") and "**" in line[2:]:
            # Bold line treated as subheading or strong text
            clean_line = line.replace("**", "")
            story.append(Paragraph(clean_line, style_subheading))
        elif line.startswith("•") or line.startswith("-"):
            # Bullet point
            story.append(Paragraph(line, style_name))
        else:
            # Normal text
            story.append(Paragraph(line, style_name))
            
    doc.build(story)
    return filename

# --- Prompts ---

VALIDATION_SYSTEM_PROMPT = """
You are an expert IT Business Analyst Hiring Manager. 
Your task is to compare a Candidate's Resume against a Job Description (JD).
You must output the result EXACTLY in the ASCII ART table format provided below. 
Do not add any markdown code blocks around the ASCII art. Just the raw text.
Ensure the boxes align correctly.
Calculate scores honestly based on keyword matching and experience relevance.
"""

VALIDATION_USER_PROMPT = """
Here is the Candidate Resume:
{resume_text}

Here is the Job Description:
{jd_text}

Generate the Validation Report in this EXACT format:

╔══════════════════════════════════════════════════════════════════════════════╗
║                 IT-BUSINESS ANALYST RESUME VALIDATION REPORT                  ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   Candidate: [Extract Name]                                                        ║
║   Job Position: [Extract Job Title]                         ║
║   Analysis Date: [Current Date]                                                  ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   🎯 **OVERALL ELIGIBILITY: [XX]%**                                              ║
║   ✅ **GOOD MATCH - RECOMMENDED** (Or ❌ **LOW MATCH**)                                              ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   📈 **SCORE BREAKDOWN:**                                                      ║
║                                                                              ║
║   • Education:           **[XX]%**  [Draw Bar]                                ║
║   • Functional Skills:   **[XX]%**  [Draw Bar]                                 ║
║   • Experience:          **[XX]%**  [Draw Bar]                                 ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   🔍 **DETAILED ANALYSIS BY FACTOR**                                           ║
║                                                                              ║
║   🎓 **EDUCATION ([XX]% Match)**                                                 ║
║   ──────────────────────────────────────────────────────────────────────────║
║   ✅ **Matched:**                                                              ║
║      • [List Matched Items]                   ║
║                                                                              ║
║   ❌ **Missing:**                                                              ║
║      • [List Missing Items]                         ║
║                                                                              ║
║   🔧 **FUNCTIONAL SKILLS ([XX]% Match)**                                         ║
║   ──────────────────────────────────────────────────────────────────────────║
║   ✅ **Matched Skills:**                                                       ║
║      • [List Skills]                                            ║
║                                                                              ║
║   ❌ **Missing Skills:**                                                       ║
║      • [List Skills]                                        ║
║                                                                              ║
║   💼 **EXPERIENCE ([XX]% Match)**                                                ║
║   ──────────────────────────────────────────────────────────────────────────║
║   ✅ **Matched:**                                                              ║
║      • [List Matches]                           ║
║                                                                              ║
║   ❌ **Missing/Gaps:**                                                         ║
║      • [List Gaps]        ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   💡 **RECOMMENDATIONS:**                                                      ║
║                                                                              ║
║   1. [Recommendation 1]     ║
║   2. [Recommendation 2]       ║
║   3. [Recommendation 3]                     ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   📊 **FACTOR SUMMARY:**                                                       ║
║   ──────────────────────────────────────────────────────────────────────────║
║   • Education:        [Summary]       ║
║   • Skills:           [Summary]       ║
║   • Experience:       [Summary]       ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   📌 **ELIGIBILITY ASSESSMENT:**                                               ║
║   **[XX]% Overall - [Verdict]**    ║
║                                                                              ║
║   • [Key Point 1]                         ║
║   • [Key Point 2]                              ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

RESUME_GEN_SYSTEM_PROMPT = """
You are an expert Resume Writer specializing in ATS-optimized IT Business Analyst resumes.
You must rewrite the user's resume to align with the provided Job Description.
You must follow the EXACT structure and style of the provided 'Priya Sharma' ATS Template.
Structure Requirements:
1. Header: Name | Title, Contact Info (Email, Phone, Location, LinkedIn, GitHub)
2. Professional Summary
3. Domain Expertise (List format)
4. Professional Experience (Role, Company, Location, Date, Bullets)
5. Certifications
6. Technical & Professional Skills (Tools, Databases, Methodologies, Frameworks, Languages)
7. Key Projects
8. Education
9. Footer: ATS-Optimized Resume | IT Business Analyst | Last Updated: [Date]

Style Requirements:
- Single column.
- No graphics, no tables, no photos.
- Clean hierarchy.
- Use standard fonts (implied in text).
- Optimize keywords based on the Job Description.
- Do not use Markdown headers (#), use text formatting (Capitalization, Bold via **).
"""

RESUME_GEN_USER_PROMPT = """
Here is the User's Original Resume Data:
{resume_text}

Here is the Target Job Description:
{jd_text}

Here is the Reference Template Structure (Follow this layout):
[Name]
[Title]
[Contact Info]

Professional Summary
[Summary]

Domain Expertise
[List]

Professional Experience
[Role][Company] — [Location][Date]
[Bullets]

Certifications
[List]

Technical & Professional Skills
[Categories]

Key Projects
[Project Details]

Education
[Details]

Generate the new resume content now.
"""

# --- Main App Logic ---

def main():
    st.title("📊 Business Analyst Job Apply Pro")
    st.markdown("### ATS Optimized Resume Validator & Generator")
    
    # Sidebar for API Key
    with st.sidebar:
        st.header("⚙️ Configuration")
        api_key = st.text_input("OpenRouter API Key", type="password", help="Get your key from openrouter.ai")
        st.markdown("---")
        st.info("💡 **Tip:** Use a Qwen model via OpenRouter for best results.")
        st.markdown("### About")
        st.markdown("This tool validates your BA resume against a JD and generates an ATS-optimized version.")

    if not api_key:
        st.warning("Please enter your OpenRouter API Key in the sidebar to proceed.")
        st.stop()

    # Main Columns
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Upload Your Resume")
        uploaded_file = st.file_uploader("Choose a file", type=['pdf', 'docx', 'txt'])
        
        st.subheader("2. Job Requirements")
        jd_text = st.text_area("Paste Employer's Job Description", height=200, placeholder="Paste the full job description here...")

    with col2:
        st.subheader("3. Actions")
        validate_btn = st.button("🔍 Validate Resume", type="primary")
        st.markdown("---")
        st.write("### Status")
        
    # Session State Initialization
    if 'validation_report' not in st.session_state:
        st.session_state.validation_report = None
    if 'generated_resume' not in st.session_state:
        st.session_state.generated_resume = None
    if 'resume_text' not in st.session_state:
        st.session_state.resume_text = None

    # Logic
    if uploaded_file and jd_text:
        if st.session_state.resume_text is None:
            with st.spinner("Parsing Resume..."):
                st.session_state.resume_text = extract_text_from_file(uploaded_file)
        
        if validate_btn:
            with st.spinner("Analyzing Resume against JD..."):
                prompt = VALIDATION_USER_PROMPT.format(
                    resume_text=st.session_state.resume_text,
                    jd_text=jd_text
                )
                report = call_openrouter_api(prompt, VALIDATION_SYSTEM_PROMPT, api_key)
                if report:
                    st.session_state.validation_report = report
                    st.rerun()
        
        if st.session_state.validation_report:
            st.success("Validation Complete!")
            st.markdown("### 📋 Validation Report")
            # Display ASCII art in a code block to preserve formatting
            st.code(st.session_state.validation_report, language="text")
            
            st.markdown("---")
            st.subheader("4. Update Resume?")
            update_choice = st.radio("Do you want to update your resume as per employer's latest job's profile?", ["Yes", "No"], horizontal=True)
            
            if update_choice == "Yes":
                if st.button("✨ Generate ATS Optimized Resume"):
                    with st.spinner("Generating New Resume based on Priya Sharma Template..."):
                        prompt = RESUME_GEN_USER_PROMPT.format(
                            resume_text=st.session_state.resume_text,
                            jd_text=jd_text
                        )
                        new_resume = call_openrouter_api(prompt, RESUME_GEN_SYSTEM_PROMPT, api_key)
                        if new_resume:
                            st.session_state.generated_resume = new_resume
                            st.rerun()
    
    elif uploaded_file and not jd_text:
        st.info("Please paste the Job Description to proceed.")
    elif jd_text and not uploaded_file:
        st.info("Please upload your Resume to proceed.")

    # Display Generated Resume
    if st.session_state.generated_resume:
        st.markdown("---")
        st.subheader("📄 Generated ATS Resume")
        st.text_area("Preview", value=st.session_state.generated_resume, height=400)
        
        # Generate PDF
        pdf_filename = "ATS_Optimized_Resume.pdf"
        generate_pdf_resume(st.session_state.generated_resume, pdf_filename)
        
        with open(pdf_filename, "rb") as pdf_file:
            st.download_button(
                label="📥 Download PDF (A4 Format)",
                data=pdf_file,
                file_name=pdf_filename,
                mime="application/pdf"
            )
        
        # Cleanup temp file
        if os.path.exists(pdf_filename):
            os.remove(pdf_filename)

if __name__ == "__main__":
    main()
