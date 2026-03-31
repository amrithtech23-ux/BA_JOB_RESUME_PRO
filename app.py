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
from datetime import datetime

# Page Configuration
st.set_page_config(
    page_title="Business Analyst Job Apply Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better ASCII rendering and UI
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
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .warning-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        color: #856404;
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
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
            return text
        elif uploaded_file.name.endswith('.docx'):
            doc = Document(uploaded_file)
            return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        elif uploaded_file.name.endswith('.txt'):
            return uploaded_file.read().decode("utf-8")
        else:
            st.error("❌ Unsupported file format. Please upload PDF, DOCX, or TXT.")
            return None
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        return None

def call_openrouter_api(prompt, system_instruction, api_key, model="qwen/qwen-2.5-72b-instruct"):
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
        "max_tokens": 4000
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content']
    except requests.exceptions.Timeout:
        st.error("⏱️ Request timed out. The API is taking too long. Please try again.")
        return None
    except requests.exceptions.HTTPError as e:
        st.error(f"❌ HTTP Error: {e.response.status_code} - {e.response.text}")
        return None
    except Exception as e:
        st.error(f"❌ API Error: {str(e)}")
        return None

def generate_pdf_resume(resume_text, filename="ATS_Optimized_Resume.pdf"):
    """Generates a simple ATS-friendly PDF based on the text content."""
    try:
        doc = SimpleDocTemplate(
            filename, 
            pagesize=A4, 
            rightMargin=0.75*inch, 
            leftMargin=0.75*inch, 
            topMargin=0.75*inch, 
            bottomMargin=0.75*inch
        )
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
        
        style_contact = ParagraphStyle(
            'CustomContact',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            leading=11,
            alignment=TA_LEFT,
            spaceAfter=4
        )

        story = []
        
        # Simple parsing logic to format the generated text into PDF elements
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
            elif line.startswith("•") or line.startswith("-") or line.startswith("📧") or line.startswith("📞"):
                # Bullet point or contact info
                story.append(Paragraph(line, style_name))
            else:
                # Normal text
                story.append(Paragraph(line, style_name))
                
        doc.build(story)
        return filename
    except Exception as e:
        st.error(f"❌ Error generating PDF: {str(e)}")
        return None

# --- Prompts ---

VALIDATION_SYSTEM_PROMPT = """
You are an expert IT Business Analyst Hiring Manager with 15+ years of experience.
Your task is to compare a Candidate's Resume against a Job Description (JD).
You must output the result EXACTLY in the ASCII ART table format provided below.
Do not add any markdown code blocks around the ASCII art. Just the raw text.
Ensure the boxes align correctly using box-drawing characters.
Calculate scores honestly based on keyword matching, experience relevance, and skill alignment.
Be specific about what's matched and what's missing.
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
║   • Education:           **[XX]%**  [Draw Bar with █ and ▒]                                ║
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
You must rewrite the user's resume to align with the provided Job Description.
You must follow the EXACT structure and style of the provided 'Priya Sharma' ATS Template.

Structure Requirements (FOLLOW THIS EXACT ORDER):
1. Name (just the name, centered)
2. Title (IT Business Analyst | Agile BA | Requirements Engineer)
3. Contact Info (Email, Phone, Location, LinkedIn, GitHub)
4. Professional Summary (3-4 lines)
5. Domain Expertise (list format, no bullets)
6. Professional Experience (Role, Company — Location, Date, Bullets)
7. Certifications (with emojis 🏅)
8. Technical & Professional Skills (Tools, Databases, Methodologies, Frameworks, Languages)
9. Key Projects (with 📌 emoji)
10. Education (Degree — Institution | Year | Grade)
11. Footer: ATS-Optimized Resume | IT Business Analyst | Last Updated: [Current Month Year]

Style Requirements:
- Single column layout
- No graphics, no tables, no photos
- Clean hierarchy with clear section headers
- Use standard fonts (Helvetica/Arial implied)
- Optimize keywords based on the Job Description
- Do not use Markdown headers (#), use plain text with capitalization
- Use emojis sparingly (🏅 for certs, 📌 for projects, 📧📞🔗 for contact)
- Keep it concise and achievement-oriented
- Quantify achievements where possible
"""

RESUME_GEN_USER_PROMPT = """
Here is the User's Original Resume Data:
{resume_text}

Here is the Target Job Description:
{jd_text}

Here is the Reference Template Structure (Follow this EXACT layout):

Priya Sharma
IT Business Analyst | Agile BA | Requirements Engineer
📧 priya.sharma@example.com | 📞 +91 98765 43210 | 📍 Bengaluru, India
🔗 linkedin.com/in/priyasharma | 💻 github.com/priyaba

Professional Summary
[3-4 lines summary]

Domain Expertise
[List domains without bullets]

Professional Experience
[Role][Company] — [Location][Date]
[Bullet achievements]

Certifications
🏅 [Cert Name] — [Issuer] | [Year]

Technical & Professional Skills
Tools: [list]
Databases: [list]
Methodologies: [list]
Frameworks: [list]
Languages: [list]

Key Projects
📌 [Project Name]
[Description]

Education
[Degree] — [Institution] | [Year] | [Grade]

ATS-Optimized Resume | IT Business Analyst | Last Updated: [Month Year]

Generate the new resume content now following this exact structure.
"""

# --- Main App Logic ---

def main():
    st.title("📊 Business Analyst Job Apply Pro")
    st.markdown("### ATS Optimized Resume Validator & Generator")
    st.markdown("---")
    
    # Sidebar for API Key and Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Try to get API key from secrets first, then from user input
        if "OPENROUTER_API_KEY" in st.secrets:
            api_key = st.secrets["OPENROUTER_API_KEY"]
            st.success("✅ API Key loaded from secrets")
            st.info("💡 Using secure environment configuration")
        else:
            api_key = st.text_input(
                "OpenRouter API Key", 
                type="password",
                help="Get your key from https://openrouter.ai/keys",
                placeholder="sk-or-v1-..."
            )
            if api_key:
                st.success("✅ API Key entered")
            else:
                st.warning("⚠️ Please enter your OpenRouter API Key")
                st.markdown("[Get your API key here](https://openrouter.ai/keys)")
        
        st.markdown("---")
        st.markdown("### 📋 Model Settings")
        model = st.selectbox(
            "AI Model",
            ["qwen/qwen-2.5-72b-instruct", "qwen/qwen-2.5-coder-32b-instruct"],
            index=0,
            help="Qwen model for best results"
        )
        
        st.markdown("---")
        st.markdown("### ℹ️ About")
        st.markdown("""
        **Version:** 1.0.0
        
        This tool validates your BA resume against a JD and generates an ATS-optimized version.
        
        **Powered by:**
        - Streamlit
        - OpenRouter API
        - Qwen Models
        """)
        
        st.markdown("---")
        st.markdown("### 🔗 Quick Links")
        st.markdown("[GitHub Repository](https://github.com/amrithtech23-ux/BA_JOB_RESUME_PRO)")
        st.markdown("[OpenRouter](https://openrouter.ai/)")

    # Check if API key is available
    if not api_key:
        st.warning("⚠️ **Please enter your OpenRouter API Key in the sidebar to proceed.**")
        st.markdown("""
        ### 🚀 Getting Started:
        1. Get your API key from [OpenRouter](https://openrouter.ai/keys)
        2. Enter it in the sidebar
        3. Upload your resume and paste the job description
        4. Click Validate Resume!
        """)
        st.stop()

    # Main Content Area - Two Columns
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
            placeholder="Paste the full job description here including:\n- Job title\n- Required qualifications\n- Skills and experience\n- Responsibilities\n- Any other requirements...",
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
                            model
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

if __name__ == "__main__":
    main()
