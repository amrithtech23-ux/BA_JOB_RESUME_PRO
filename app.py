import streamlit as st
import requests
import os
import tempfile
from datetime import datetime
from pypdf import PdfReader
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import inch

# ============================================
# Page Configuration
# ============================================
st.set_page_config(
    page_title="Business Analyst Job Apply Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# Custom CSS
# ============================================
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
        font-size: 11px;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
        margin: 1rem 0;
    }
    .stButton>button {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================
# Helper Functions
# ============================================

def extract_text_from_file(uploaded_file):
    """Extracts text from PDF, DOCX, or TXT files."""
    try:
        text = ""
        if uploaded_file.name.endswith('.pdf'):
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        elif uploaded_file.name.endswith('.docx'):
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                if para.text.strip():
                    text += para.text + "\n"
        elif uploaded_file.name.endswith('.txt'):
            text = uploaded_file.read().decode("utf-8")
        else:
            return None
        return text.strip()
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        return None

def get_api_key():
    """Get API key from secrets or user input."""
    try:
        if hasattr(st, 'secrets') and "OPENROUTER_API_KEY" in st.secrets:
            return st.secrets["OPENROUTER_API_KEY"]
    except:
        pass
    return st.text_input("OpenRouter API Key", type="password", 
                        help="Get your key from https://openrouter.ai/keys")

def call_openrouter_api(prompt, system_instruction, api_key, model="qwen/qwen-2.5-72b-instruct", max_tokens=4000):
    """Calls OpenRouter API using Qwen model."""
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/amrithtech23-ux/BA_JOB_RESUME_PRO",
        "X-Title": "BA Job Apply Pro"
    }
    
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3,
        "max_tokens": max_tokens
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content']
    except requests.exceptions.Timeout:
        st.error("⏱️ Request timed out. Please try again.")
        return None
    except requests.exceptions.HTTPError as e:
        st.error(f"❌ HTTP Error: {e.response.status_code}")
        return None
    except Exception as e:
        st.error(f"❌ API Error: {str(e)}")
        return None

def generate_pdf_resume(resume_text, filename="ATS_Optimized_Resume.pdf"):
    """Generates ATS-friendly PDF following Priya Sharma template structure."""
    try:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_filename = temp_file.name
        temp_file.close()
        
        doc = SimpleDocTemplate(
            temp_filename, 
            pagesize=A4,
            rightMargin=0.75*inch, 
            leftMargin=0.75*inch,
            topMargin=0.75*inch, 
            bottomMargin=0.75*inch
        )
        
        styles = getSampleStyleSheet()
        
        # Custom Styles - ATS Optimized (Single column, standard fonts)
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
            alignment=TA_CENTER
        )
        
        style_subheading = ParagraphStyle(
            'CustomSubHeading',
            parent=styles['Heading3'],
            fontName='Helvetica-Bold',
            fontSize=11,
            leading=13,
            spaceBefore=8,
            spaceAfter=4,
            alignment=TA_LEFT
        )
        
        style_contact = ParagraphStyle(
            'CustomContact',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            leading=11,
            alignment=TA_CENTER,
            spaceAfter=4
        )
        
        style_footer = ParagraphStyle(
            'CustomFooter',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            leading=10,
            alignment=TA_CENTER,
            textColor='#666666'
        )

        story = []
        lines = resume_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 4))
                continue
            
            # Detect formatting based on content patterns
            if line.startswith("###"):
                story.append(Paragraph(line.replace("###", "").strip(), style_heading))
            elif line.startswith("**") and "**" in line[2:]:
                clean_line = line.replace("**", "")
                story.append(Paragraph(clean_line, style_subheading))
            elif any(line.startswith(prefix) for prefix in ["📧", "", "📍", "🔗", "💻"]):
                story.append(Paragraph(line, style_contact))
            elif line.startswith("•") or line.startswith("-") or line.startswith("🏅") or line.startswith("📌"):
                story.append(Paragraph(line, style_name))
            elif "ATS-Optimized Resume" in line or "Last Updated" in line:
                story.append(Paragraph(line, style_footer))
            else:
                story.append(Paragraph(line, style_name))
        
        doc.build(story)
        return temp_filename
    except Exception as e:
        st.error(f"❌ Error generating PDF: {str(e)}")
        return None

# ============================================
# Prompts
# ============================================

VALIDATION_SYSTEM_PROMPT = """
You are an expert IT Business Analyst Hiring Manager with 15+ years of experience.
Your task is to compare a Candidate's Resume against a Job Description (JD).

OUTPUT REQUIREMENTS:
1. Output the result EXACTLY in the ASCII ART table format provided
2. Do NOT add any markdown code blocks around the ASCII art
3. Ensure the boxes align correctly using box-drawing characters (╔═╗║╚═╝╠╣)
4. Calculate scores honestly based on keyword matching, experience relevance, and skill alignment
5. Be specific about what's matched and what's missing
6. Use the exact section headers as shown in the template
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
║   Analysis Date: [Current Date in YYYY-MM-DD format]                                                  ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   🎯 **OVERALL ELIGIBILITY: [XX]%**                                              ║
║   ✅ **GOOD MATCH - RECOMMENDED** (Or ❌ **LOW MATCH - NOT RECOMMENDED**)                                              ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   📈 **SCORE BREAKDOWN:**                                                      ║
║                                                                              ║
║   • Education:           **[XX]%**  [Draw Bar with █ and ▒ - 10 chars]                                ║
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
║      • [List Matched Items - be specific]                   ║
║                                                                              ║
║   ❌ **Missing:**                                                              ║
║      • [List Missing Items]                         ║
║                                                                              ║
║   🔧 **FUNCTIONAL SKILLS ([XX]% Match)**                                         ║
║   ──────────────────────────────────────────────────────────────────────────║
║   ✅ **Matched Skills:**                                                       ║
║      • [List Skills - be specific]                                            ║
║                                                                              ║
║   ❌ **Missing Skills:**                                                       ║
║      • [List Skills - be specific]                                        ║
║                                                                              ║
║   💼 **EXPERIENCE ([XX]% Match)**                                                ║
║   ──────────────────────────────────────────────────────────────────────────║
║   ✅ **Matched:**                                                              ║
║      • [List Matches - be specific]                           ║
║                                                                              ║
║   ❌ **Missing/Gaps:**                                                         ║
║      • [List Gaps - be specific]        ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   💡 **RECOMMENDATIONS:**                                                      ║
║                                                                              ║
║   1. [Specific Recommendation 1]     ║
║   2. [Specific Recommendation 2]       ║
║   3. [Specific Recommendation 3]                     ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   📊 **FACTOR SUMMARY:**                                                       ║
║   ──────────────────────────────────────────────────────────────────────────║
║   • Education:        [Brief Summary]       ║
║   • Skills:           [Brief Summary]       ║
║   • Experience:       [Brief Summary]       ║
║                                                                              ║
╠══════════════════════════════════════════════════════════════════════════════╣
║                                                                              ║
║   📌 **ELIGIBILITY ASSESSMENT:**                                               ║
║   **[XX]% Overall - [Verdict: RECOMMENDED/NOT RECOMMENDED for interview]**    ║
║                                                                              ║
║   • [Key Point 1]                         ║
║   • [Key Point 2]                              ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

RESUME_GEN_SYSTEM_PROMPT = """
You are an expert Resume Writer specializing in ATS-optimized IT Business Analyst resumes.

TASK: Rewrite the user's resume to align with the provided Job Description.

TEMPLATE STRUCTURE (FOLLOW EXACTLY - Based on Priya Sharma ATS Template):

1. Name (just the name, centered)
2. Title (IT Business Analyst | Agile BA | Requirements Engineer)
3. Contact Info (📧 Email | 📞 Phone | 📍 Location |  LinkedIn | 💻 GitHub)
4. Professional Summary (3-4 lines, achievement-oriented)
5. Domain Expertise (list format, no bullets, space-separated)
6. Professional Experience (Role | Company — Location | Date | Bullets with quantified achievements)
7. Certifications (🏅 Cert Name — Issuer | Year)
8. Technical & Professional Skills (Tools, Databases, Methodologies, Frameworks, Languages)
9. Key Projects (📌 Project Name | Description | Role | Tools)
10. Education (Degree — Institution | Year | Grade)
11. Footer: ATS-Optimized Resume | IT Business Analyst | Last Updated: [Month Year]

STYLE REQUIREMENTS:
- Single column layout (NO tables, NO graphics, NO photos)
- Clean hierarchy with clear section headers
- Use standard fonts (Helvetica/Arial implied)
- Optimize keywords based on the Job Description
- Do NOT use Markdown headers (#), use plain text with capitalization
- Use emojis sparingly (🏅 for certs, 📌 for projects, 📧📞 for contact)
- Keep it concise and achievement-oriented
- Quantify achievements where possible (%, $, numbers)
- Compatible with Workday, Taleo, Lever, Greenhouse ATS systems

IMPORTANT:
- Match keywords from the Job Description
- Highlight relevant BA skills (Requirements Gathering, Stakeholder Management, Agile/Scrum, etc.)
- Include domain expertise if mentioned in JD (Banking, Finance, Healthcare, etc.)
"""

RESUME_GEN_USER_PROMPT = """
Here is the User's Original Resume Data:
{resume_text}

Here is the Target Job Description:
{jd_text}

Generate the new ATS-optimized resume content following the Priya Sharma template structure exactly.
"""

# ============================================
# Main Application
# ============================================

def main():
    # Title and Header
    st.title("📊 Business Analyst Job Apply Pro")
    st.markdown("### ATS Optimized Resume Validator & Generator")
    st.markdown("---")
    
    # Sidebar Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # API Key
        api_key = get_api_key()
        
        if not api_key:
            st.warning("⚠️ Please enter your OpenRouter API Key")
            st.markdown("[Get your API key here](https://openrouter.ai/keys)")
            st.stop()
        else:
            st.success("✅ API Key configured")
        
        # Model Selection
        st.markdown("---")
        st.markdown("### 🤖 Model Settings")
        model = st.selectbox(
            "AI Model",
            ["qwen/qwen-2.5-72b-instruct", "qwen/qwen-2.5-coder-32b-instruct"],
            index=0,
            help="Qwen model for best results"
        )
        
        # About Section
        st.markdown("---")
        st.markdown("### ℹ️ About")
        st.markdown("""
        **Version:** 1.0.0
        
        This tool validates your BA resume against a JD and generates an ATS-optimized version.
        
        **Powered by:**
        - Streamlit
        - OpenRouter API
        - Qwen Models
        
        **Template:** Priya Sharma ATS Format
        """)
        
        # Quick Links
        st.markdown("---")
        st.markdown("### 🔗 Links")
        st.markdown("[GitHub Repository](https://github.com/amrithtech23-ux/BA_JOB_RESUME_PRO)")
        st.markdown("[OpenRouter](https://openrouter.ai/)")
    
    # Main Content - Two Columns
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1️⃣ Upload Your Resume")
        uploaded_file = st.file_uploader(
            "Choose a file", 
            type=['pdf', 'docx', 'txt'],
            help="Upload your current resume (PDF, DOCX, or TXT)"
        )
        
        if uploaded_file:
            st.success(f"✅ Uploaded: {uploaded_file.name}")
            st.info(f"📄 Size: {uploaded_file.size / 1024:.1f} KB")
        
        st.markdown("---")
        st.subheader("2️⃣ Job Requirements")
        jd_text = st.text_area(
            "Paste Employer's Job Description", 
            height=250, 
            placeholder="""Paste the full job description including:
- Job title
- Required qualifications
- Skills and experience
- Responsibilities
- Any other requirements...""",
            help="Copy and paste the complete job posting for best results"
        )
        
        if jd_text:
            st.success(f"✅ Job description loaded ({len(jd_text)} characters)")
    
    with col2:
        st.subheader("3️⃣ Actions")
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            validate_btn = st.button("🔍 Validate Resume", type="primary", use_container_width=True)
        with col_btn2:
            reset_btn = st.button("🔄 Reset", use_container_width=True)
        
        # Reset functionality
        if reset_btn:
            st.session_state.validation_report = None
            st.session_state.generated_resume = None
            st.session_state.resume_text = None
            st.rerun()
        
        st.markdown("---")
        st.subheader("📊 Status")
        
        if 'processing' not in st.session_state:
            st.session_state.processing = False
    
    # Session State Initialization
    if 'validation_report' not in st.session_state:
        st.session_state.validation_report = None
    if 'generated_resume' not in st.session_state:
        st.session_state.generated_resume = None
    if 'resume_text' not in st.session_state:
        st.session_state.resume_text = None
    
    # Main Processing Logic
    if uploaded_file and jd_text:
        # Extract resume text if not already done
        if st.session_state.resume_text is None:
            with st.spinner("📄 Parsing resume..."):
                st.session_state.resume_text = extract_text_from_file(uploaded_file)
                if st.session_state.resume_text:
                    st.success("✅ Resume parsed successfully!")
        
        # Validation Button
        if validate_btn:
            if not st.session_state.resume_text:
                st.error("❌ Could not parse resume. Please try uploading again.")
            else:
                with st.spinner("🔍 Analyzing resume against job description..."):
                    st.session_state.processing = True
                    
                    prompt = VALIDATION_USER_PROMPT.format(
                        resume_text=st.session_state.resume_text,
                        jd_text=jd_text
                    )
                    
                    report = call_openrouter_api(
                        prompt, 
                        VALIDATION_SYSTEM_PROMPT, 
                        api_key,
                        model
                    )
                    
                    st.session_state.processing = False
                    
                    if report:
                        st.session_state.validation_report = report
                        st.success("✅ Validation complete!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to generate validation report. Please check your API key and try again.")
        
        # Display Validation Report
        if st.session_state.validation_report:
            st.markdown("---")
            st.subheader("📋 Validation Report")
            
            # Display ASCII art in a code block to preserve formatting
            st.code(st.session_state.validation_report, language="text")
            
            st.markdown("---")
            st.subheader("4️⃣ Update Resume?")
            
            update_choice = st.radio(
                "Do you want to update your resume as per employer's latest job profile?", 
                ["Yes, generate optimized resume", "No, I'm good with current resume"], 
                horizontal=True
            )
            
            if update_choice == "Yes, generate optimized resume":
                if st.button("✨ Generate ATS Optimized Resume", type="primary"):
                    with st.spinner("✍️ Generating new ATS-optimized resume..."):
                        prompt = RESUME_GEN_USER_PROMPT.format(
                            resume_text=st.session_state.resume_text,
                            jd_text=jd_text
                        )
                        
                        new_resume = call_openrouter_api(
                            prompt, 
                            RESUME_GEN_SYSTEM_PROMPT, 
                            api_key,
                            model,
                            max_tokens=5000
                        )
                        
                        if new_resume:
                            st.session_state.generated_resume = new_resume
                            st.success("✅ Resume generated successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Failed to generate resume. Please try again.")
    
    elif uploaded_file and not jd_text:
        st.info("📝 **Please paste the Job Description to proceed.**")
    elif jd_text and not uploaded_file:
        st.info("📄 **Please upload your Resume to proceed.**")
    elif not uploaded_file and not jd_text:
        st.info("👆 **Upload your resume and paste the job description to get started.**")
    
    # Display Generated Resume
    if st.session_state.generated_resume:
        st.markdown("---")
        st.subheader("📄 Generated ATS Resume")
        
        # Show preview in expandable section
        with st.expander("👁️ Preview Resume Text", expanded=True):
            st.text_area("Resume Preview", value=st.session_state.generated_resume, height=500)
        
        # Generate PDF
        pdf_filename = "ATS_Optimized_Resume.pdf"
        
        with st.spinner("📄 Creating PDF..."):
            pdf_file = generate_pdf_resume(st.session_state.generated_resume, pdf_filename)
            
            if pdf_file and os.path.exists(pdf_file):
                st.success("✅ PDF generated successfully!")
                
                # Read the PDF file for download
                with open(pdf_file, "rb") as f:
                    pdf_bytes = f.read()
                
                st.download_button(
                    label="📥 Download PDF (A4 Format)",
                    data=pdf_bytes,
                    file_name=pdf_filename,
                    mime="application/pdf",
                    type="primary"
                )
                
                # Cleanup temp file
                try:
                    os.remove(pdf_file)
                except:
                    pass
            else:
                st.error("❌ Failed to generate PDF. You can copy the text above and format it manually.")
        
        st.markdown("---")
        st.success("🎉 **Your ATS-optimized resume is ready! Good luck with your application!**")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; font-size: 12px;">
        <p>Built with ❤️ for Business Analysts worldwide | Powered by Streamlit & OpenRouter</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"❌ Application error: {str(e)}")
        st.exception(e)
