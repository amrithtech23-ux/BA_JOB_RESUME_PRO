import streamlit as st
import requests
import os
import tempfile
import re
from datetime import datetime
from pypdf import PdfReader
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.units import inch
from reportlab.lib.colors import black, gray

# ============================================
# Page Configuration - Sidebar Collapsed
# ============================================
st.set_page_config(
    page_title="Business Analyst Job Apply Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================
# Custom CSS - Hide Sidebar
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
    [data-testid="stSidebar"] {
        display: none;
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
    """Calls OpenRouter API using Qwen model with robust error handling."""
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
        
        if response.status_code != 200:
            st.error(f"❌ API Error ({response.status_code}): {response.text[:500]}")
            return None
        
        result = response.json()
        
        if 'choices' not in result:
            st.error(f"❌ Unexpected API response: {result}")
            return None
        
        if not result['choices']:
            st.error("❌ API returned empty choices")
            return None
            
        return result['choices'][0]['message']['content']
        
    except requests.exceptions.Timeout:
        st.error("⏱️ Request timed out. Please try again.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Connection error: {str(e)}")
        return None
    except Exception as e:
        st.error(f"❌ Unexpected error: {str(e)}")
        return None

def clean_resume_text(text):
    """Clean and normalize resume text before PDF generation."""
    if not text:
        return ""
    
    # Remove markdown formatting
    text = text.replace('**', '').replace('##', '').replace('#', '').replace('###', '')
    
    # Fix double commas and spacing issues
    text = re.sub(r',,+', ',', text)  # Replace multiple commas with single
    text = re.sub(r'\s+,', ',', text)  # Remove space before comma
    text = re.sub(r',\s+', ', ', text)  # Ensure single space after comma
    
    # Fix double spaces
    text = re.sub(r' {2,}', ' ', text)
    
    # Normalize bullet points
    text = text.replace('■', '•').replace('–', '-').replace('—', '|')
    
    # Remove multiple dashes used as separators
    text = re.sub(r'\n---+\n', '\n', text)
    text = re.sub(r'^---+\s*$', '', text, flags=re.MULTILINE)
    
    # Remove lines with only special characters
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        stripped = line.strip()
        # Skip lines that are just bullets, dashes, or very short
        if stripped and len(stripped) > 2 and not re.match(r'^[•\-–—\|]+$', stripped):
            cleaned_lines.append(line)
    
    text = '\n'.join(cleaned_lines)
    
    # Remove extra blank lines (more than 2 consecutive)
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text.strip()

def generate_pdf_resume(resume_text, filename="ATS_Optimized_Resume.pdf"):
    """
    Generates ATS-friendly PDF with proper formatting.
    - Professional Summary: Justified alignment
    - Domain Expertise: Left-aligned, comma-separated (clean formatting)
    - Experience/Skills/Projects/Education: Left-aligned with proper indentation
    - NO Footer, NO meta-text, NO separators
    """
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
        
        # ============================================
        # Custom Styles - Clean Professional Format
        # ============================================
        
        # Name/Title
        style_title = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=16,
            leading=18,
            alignment=TA_CENTER,
            spaceAfter=4
        )
        
        style_subtitle = ParagraphStyle(
            'SubtitleStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=12,
            alignment=TA_CENTER,
            spaceAfter=8
        )
        
        # Contact info - Centered
        style_contact = ParagraphStyle(
            'ContactStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            leading=12,
            alignment=TA_CENTER,
            spaceAfter=10
        )
        
        # Section headers - Left aligned, bold, uppercase
        style_heading = ParagraphStyle(
            'HeadingStyle',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=11,
            leading=14,
            spaceBefore=14,
            spaceAfter=6,
            alignment=TA_LEFT,
            textColor=black,
            textTransform='uppercase'
        )
        
        # Professional Summary - Justified
        style_summary = ParagraphStyle(
            'SummaryStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=13,
            alignment=TA_JUSTIFY,
            spaceAfter=10,
            leftIndent=0,
            rightIndent=0
        )
        
        # Domain Expertise - Left aligned, comma-separated
        style_domain = ParagraphStyle(
            'DomainStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=14,
            alignment=TA_LEFT,
            spaceAfter=10,
            leftIndent=0,
            rightIndent=0
        )
        
        # Content - Left aligned
        style_content = ParagraphStyle(
            'ContentStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=13,
            alignment=TA_LEFT,
            spaceAfter=5,
            leftIndent=0
        )
        
        # Bullet points - Consistent indentation
        style_bullet = ParagraphStyle(
            'BulletStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=13,
            alignment=TA_LEFT,
            spaceAfter=4,
            leftIndent=20,
            firstLineIndent=-15
        )

        # ============================================
        # Parse and Format Resume Content
        # ============================================
        story = []
        
        # Clean the text first
        resume_text = clean_resume_text(resume_text)
        lines = resume_text.split('\n')
        
        current_section = None
        is_first_line = True
        skip_section = False
        
        # Keywords to skip (meta-text, footers, separators)
        skip_keywords = [
            'critical success factors',
            'achieve 85%+ keyword match',
            'include all jd-required domains',
            'use jd language in professional summary',
            'this resume is optimized for ats',
            'designed to achieve an 85%+',
            'include, all, jd-required, skills',
            'footer:',
            'footer',
            'ats-optimized resume | it business analyst | last updated',
            'last updated:',
            'optimization notes',
            'ai-generated',
            'meta-instructional',
            'functional skills:'  # Skip the all-caps functional skills header
        ]
        
        for line in lines:
            line_stripped = line.strip()
            
            if not line_stripped:
                story.append(Spacer(1, 4))
                continue
            
            line_clean = line_stripped
            line_lower = line_clean.lower()
            
            # ============================================
            # SKIP META-INSTRUCTIONAL SECTIONS
            # ============================================
            if any(keyword in line_lower for keyword in skip_keywords):
                skip_section = True
                continue
            
            # Reset skip flag if we hit a real section header
            if any(header in line_lower for header in ['professional summary', 'domain expertise', 
                                                         'professional experience', 'certifications',
                                                         'technical', 'key projects', 'education']):
                skip_section = False
            
            if skip_section:
                continue
            
            # ============================================
            # Detect Section Headers
            # ============================================
            section_keywords = {
                'professional summary': 'professional summary',
                'domain expertise': 'domain expertise',
                'professional experience': 'professional experience',
                'certifications': 'certifications',
                'technical': 'technical & professional skills',
                'key projects': 'key projects',
                'education': 'education'
            }
            
            is_header = False
            for keyword, section_name in section_keywords.items():
                if keyword in line_lower:
                    story.append(Spacer(1, 8))
                    story.append(Paragraph(line_clean.upper(), style_heading))
                    story.append(Spacer(1, 2))
                    current_section = section_name
                    is_header = True
                    break
            
            if is_header:
                continue
            
            # ============================================
            # Format Content by Section
            # ============================================
            
            # Name (first non-empty line)
            if is_first_line and len(line_clean.split()) <= 3:
                if not any(c in line_clean for c in ['@', '📧', '📞', '📍', '■', '|', '—', '•', ':', '+']):
                    story.append(Paragraph(line_clean, style_title))
                    is_first_line = False
                    continue
            
            # Title/Role
            if is_first_line and any(term in line_clean for term in ['Business Analyst', 'Agile BA', 'Requirements Engineer']):
                story.append(Paragraph(line_clean, style_subtitle))
                is_first_line = False
                continue
            
            # Contact info
            if any(icon in line_clean for icon in ['📧', '📞', '📍', '■', '@', '+91', 'linkedin', 'github', 'email']):
                if '|' in line_clean or '@' in line_clean or '+' in line_clean:
                    story.append(Paragraph(line_clean, style_contact))
                    continue
            
            # Professional Summary - Justified
            if current_section == 'professional summary':
                if not line_clean.startswith('•') and not line_clean.startswith('-'):
                    story.append(Paragraph(line_clean, style_summary))
                continue
            
            # Domain Expertise - Comma-separated with clean formatting
            if current_section == 'domain expertise':
                # Clean up: remove double commas, extra spaces, normalize
                domains = [d.strip() for d in line_clean.replace('  ', ' ').split(',') if d.strip() and len(d.strip()) > 2]
                # Rejoin with proper comma spacing
                domain_line = ', '.join(domains)
                if domain_line and len(domain_line) > 5:  # Only add if meaningful content
                    story.append(Paragraph(domain_line, style_domain))
                continue
            
            # Professional Experience - Left aligned with bullets
            if current_section == 'professional experience':
                if '|' in line_clean and '—' in line_clean:
                    parts = line_clean.split('|')
                    if len(parts) >= 2:
                        role = parts[0].strip()
                        rest = '|'.join(parts[1:]).strip()
                        formatted = f"<b>{role}</b> | {rest}"
                        story.append(Paragraph(formatted, style_content))
                    else:
                        story.append(Paragraph(line_clean, style_content))
                elif line_clean.startswith('•') or line_clean.startswith('-'):
                    content = line_clean[1:].strip()
                    if content and len(content) > 10:  # Only add if there's actual content
                        bullet_text = f"• {content}"
                        story.append(Paragraph(bullet_text, style_bullet))
                else:
                    story.append(Paragraph(line_clean, style_content))
                continue
            
            # Education - Left aligned
            if current_section == 'education':
                if '—' in line_clean or '|' in line_clean:
                    story.append(Paragraph(line_clean, style_content))
                elif line_clean.startswith('•') or line_clean.startswith('-'):
                    content = line_clean[1:].strip()
                    if content and len(content) > 10:
                        story.append(Paragraph(f"• {content}", style_bullet))
                else:
                    story.append(Paragraph(line_clean, style_content))
                continue
            
            # Certifications - Left aligned with bullets
            if current_section == 'certifications':
                if line_clean.startswith('•') or line_clean.startswith('-') or line_clean.startswith('■') or line_clean.startswith('🏅'):
                    content = line_clean[1:].strip()
                    if content and len(content) > 10:
                        story.append(Paragraph(f"• {content}", style_bullet))
                else:
                    story.append(Paragraph(line_clean, style_content))
                continue
            
            # Technical Skills - Left aligned
            if current_section == 'technical & professional skills':
                if line_clean.startswith('•') or line_clean.startswith('-') or line_clean.startswith('**'):
                    content = line_clean.replace('**', '').strip()
                    if content.startswith('-') or content.startswith('•'):
                        content = content[1:].strip()
                    if content and len(content) > 5:
                        story.append(Paragraph(f"• {content}", style_bullet))
                else:
                    story.append(Paragraph(line_clean, style_content))
                continue
            
            # Key Projects - Left aligned with bullets
            if current_section == 'key projects':
                if line_clean.startswith('•') or line_clean.startswith('-') or line_clean.startswith('■') or line_clean.startswith('📌'):
                    content = line_clean[1:].strip()
                    if content and len(content) > 10:
                        story.append(Paragraph(f"• {content}", style_bullet))
                else:
                    story.append(Paragraph(line_clean, style_content))
                continue
            
            # Default fallback
            story.append(Paragraph(line_clean, style_content))
        
        # ============================================
        # Build PDF
        # ============================================
        doc.build(story)
        return temp_filename
        
    except Exception as e:
        st.error(f"❌ Error generating PDF: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

# ============================================
# Prompts - Strategic Skill Matching for 85%+ Match
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
7. List ALL missing skills/domains that should be added to resume for 85%+ match
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

CRITICAL OBJECTIVE:
Generate a resume that achieves **85%+ keyword match** with the Job Description.

OUTPUT FORMAT RULES (CRITICAL):
1. NO double commas (,,) anywhere in the resume
2. NO markdown formatting (**, ##, #, etc.)
3. NO separator lines (---)
4. NO footer lines
5. NO meta-instructional text
6. Use standard bullet points (•) consistently
7. Clean, professional formatting only
8. NO all-caps section headers in content (like "FUNCTIONAL SKILLS:")
9. Proper spacing between sections

STRATEGIC APPROACH FOR 85%+ MATCH:

1. **DOMAIN EXPERTISE - COMBINE CANDIDATE + JOB REQUIREMENTS:**
   - Format: "Banking, Insurance, Financial Services, Healthcare, E-commerce" (single commas, proper spacing)
   - Include ALL domains from candidate's resume PLUS ALL domains from job description
   - 5-6 domains per line for proper wrapping

2. **SKILLS INTEGRATION - MERGE CANDIDATE + JD SKILLS:**
   - Add ALL tools/skills from JD that relate to candidate's experience
   - Mark unfamiliar skills as "(Familiar)" or "(Basic)" if needed
   - Include 90%+ of JD keywords in skills section

3. **PROFESSIONAL SUMMARY - KEYWORD OPTIMIZATION:**
   - Include 8-10 keywords from job description
   - Mention years of experience (from candidate)
   - Include domain expertise (candidate's + job's domains)
   - Add quantified achievements (from candidate's resume)
   - Length: 4-5 lines

4. **EXPERIENCE BULLETS - JD KEYWORD INTEGRATION:**
   - Use exact phrases from job description
   - Frame candidate's achievements using JD language
   - Add metrics from candidate's original resume
   - Each bullet should include 1-2 JD keywords

TEMPLATE STRUCTURE (FOLLOW EXACTLY - Priya Sharma Format):

1. Name (centered, bold, 16pt)
2. Title (IT Business Analyst | Agile BA | Requirements Engineer)
3. Contact Info (📧 Email |  Phone | 📍 Location | LinkedIn | GitHub)
4. Professional Summary (4-5 lines, justified, 80% JD keywords)
5. Domain Expertise (Candidate's domains + Job's domains, comma-separated with SINGLE commas)
6. Professional Experience (Role | Company — Location | Date | Bullets)
7. Certifications (🏅 Cert Name — Issuer | Year)
8. Technical & Professional Skills (All candidate skills + JD skills)
9. Key Projects (📌 Project Name | Description | Role | Tools)
10. Education (Degree — Institution | Year | Grade)

FORMATTING RULES:
- Single column layout (NO tables, NO graphics, NO photos)
- Standard fonts (Helvetica/Arial)
- Use emojis: 📧 for contact, 🏅 for certs, 📌 for projects
- Quantify achievements (%, $, numbers)
- ATS-compatible (Workday, Taleo, Lever, Greenhouse)
- NO footer line at bottom
- NO meta-instructional text
- NO double commas (,,)
- NO separator lines (---)
- NO all-caps text in content sections
- OUTPUT ONLY THE RESUME CONTENT - no explanations, no notes

DOMAIN EXPERTISE EXAMPLE (CORRECT FORMAT):
Banking, Insurance, Financial Services, Healthcare, E-commerce, Retail

SKILLS SECTION EXAMPLE (CORRECT FORMAT):
• Tools: JIRA, Confluence, Tableau, Power BI, Qlik Sense, Visio, Lucidchart
• Databases: SQL (Advanced), PL/SQL, MySQL, PostgreSQL
• Methodologies: Agile, Scrum, SAFe, Waterfall, SDLC
• ETL Tools: Talend, Informatica, SSIS (Familiar)
• Languages: Python (Intermediate), PySpark (Basic)

CRITICAL OUTPUT RULES:
- Achieve 85%+ keyword match with JD
- Include ALL JD-required domains in Domain Expertise (with SINGLE commas)
- Include ALL JD-required skills in Skills section
- Use JD language in Professional Summary and Experience bullets
- Maintain credibility by marking unfamiliar skills appropriately
- Quantify achievements with candidate's actual metrics
- OUTPUT ONLY RESUME CONTENT - no footers, no notes, no AI comments
- NO double commas anywhere
- NO markdown formatting
- NO all-caps section headers in content
"""

RESUME_GEN_USER_PROMPT = """
Here is the User's Original Resume Data:
{resume_text}

Here is the Target Job Description:
{jd_text}

INSTRUCTIONS FOR 85%+ MATCH:

1. **EXTRACT FROM CANDIDATE:**
   - Actual experience (years, roles, companies)
   - Real education and certifications
   - Genuine achievements with metrics
   - Current skills and tools

2. **EXTRACT FROM JOB DESCRIPTION:**
   - Required domains (Banking, Insurance, Healthcare, etc.)
   - Required skills (SQL, Tableau, Agile, etc.)
   - Required tools (JIRA, Qlik Sense, ETL, etc.)
   - Required experience level

3. **CREATE OPTIMIZED RESUME:**
   
   **Professional Summary:**
   - Include candidate's years of experience
   - Add 8-10 keywords from JD
   - Mention candidate's domains + JD domains
   - Include 2-3 quantified achievements
   
   **Domain Expertise:**
   - List ALL domains from candidate's resume
   - Add ALL domains from job description
   - Format: Comma-separated with SINGLE commas (e.g., "Banking, Insurance, Financial Services")
   - 5-6 domains per line
   
   **Skills Section:**
   - Include ALL candidate's skills
   - Add ALL JD-required skills (mark as "Familiar" if not expert)
   - Group by category: Tools, Databases, Methodologies, ETL, Languages
   
   **Experience Bullets:**
   - Use candidate's actual achievements
   - Frame using JD keywords
   - Add metrics from original resume
   - Each bullet includes 1-2 JD keywords
   
   **Certifications:**
   - Keep candidate's actual certifications
   - Add relevant JD certifications as "In Progress" if appropriate

4. **KEYWORD DENSITY:**
   - Professional Summary: 8-10 JD keywords
   - Domain Expertise: 100% of JD domains
   - Skills Section: 90%+ of JD skills
   - Experience Bullets: 70%+ of JD keywords

5. **FINAL OUTPUT RULES:**
   - OUTPUT ONLY THE RESUME CONTENT
   - NO footer line at bottom
   - NO meta-instructional text
   - NO AI comments, explanations, or notes
   - NO double commas (,,)
   - NO markdown formatting (**, ##, #)
   - NO separator lines (---)
   - NO all-caps text in content sections
   - Just the clean resume following Priya Sharma template

Generate the ATS-optimized resume following this exact structure to achieve 85%+ match.
"""

# ============================================
# Main Application
# ============================================

def main():
    # Title and Header
    st.title("📊 Business Analyst Job Apply Pro")
    st.markdown("### ATS Optimized Resume Validator & Generator")
    st.markdown("---")
    
    # Sidebar Configuration (Collapsed by default)
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
        
        **Features:**
        - Resume validation with ASCII report
        - 85%+ JD match optimization
        - ATS-friendly PDF generation
        - Priya Sharma template format
        
        **Powered by:**
        - Streamlit
        - OpenRouter API
        - Qwen Models
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
- Domain requirements (Banking, Insurance, etc.)
- Tools and technologies""",
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
                ["Yes, generate optimized resume (85%+ match)", "No, I'm good with current resume"], 
                horizontal=True
            )
            
            if update_choice == "Yes, generate optimized resume (85%+ match)":
                if st.button("✨ Generate ATS Optimized Resume", type="primary"):
                    with st.spinner("✍️ Generating new ATS-optimized resume (targeting 85%+ match)..."):
                        prompt = RESUME_GEN_USER_PROMPT.format(
                            resume_text=st.session_state.resume_text,
                            jd_text=jd_text
                        )
                        
                        new_resume = call_openrouter_api(
                            prompt, 
                            RESUME_GEN_SYSTEM_PROMPT, 
                            api_key,
                            model,
                            max_tokens=6000
                        )
                        
                        if new_resume:
                            st.session_state.generated_resume = new_resume
                            st.success("✅ Resume generated successfully! Target: 85%+ JD match")
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
        st.success("🎉 **Your ATS-optimized resume is ready! Target: 85%+ JD match. Good luck with your application!**")
    
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
