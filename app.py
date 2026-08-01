import base64
import io
import json
import os
import streamlit as st
from docxtpl import DocxTemplate
from groq import Groq
from pydantic import BaseModel
from docx.oxml import OxmlElement
from pypdf import PdfReader

# ---------------------------------------------------------
# Page Config & School Color Theme
# ---------------------------------------------------------
st.set_page_config(
    page_title="Florida School of Skills - CAPS Lesson Plan Generator",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Responsive & Modern Guided CSS Framework
st.markdown("""
    <style>
    /* Global Page Styling */
    .stApp {
        background-color: #f3f4f6;
        color: #111827;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
    }
    
    /* Hide Default Sidebar Completely */
    [data-testid="stSidebar"] {
        display: none;
    }
    
    /* Main Container Centering & Max Width */
    .main .block-container {
        max-width: 1040px !important;
        padding-top: 0rem !important;
        padding-bottom: 3rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* Top Compact Header */
    .compact-header {
        background-color: #111827;
        border-bottom: 4px solid #721f2f;
        padding: 16px 24px;
        border-radius: 0 0 12px 12px;
        display: flex;
        align-items: center;
        gap: 20px;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .header-badge-img {
        width: auto !important;
        max-width: 60px !important;
        height: 55px !important;
        max-height: 55px !important;
        object-fit: contain !important;
        border-radius: 4px;
        flex-shrink: 0 !important;
    }
    .header-text-container h1 {
        color: #ffffff !important;
        font-size: 1.5rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        letter-spacing: 0.5px;
        line-height: 1.2;
    }
    .header-text-container p {
        color: #9ca3af !important;
        font-size: 0.85rem !important;
        margin-top: 2px !important;
        margin-bottom: 0 !important;
        font-weight: 600;
        letter-spacing: 0.4px;
    }

    /* Step Cards Base Styling */
    .step-card {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    .step-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 16px;
        border-bottom: 2px solid #f3f4f6;
        padding-bottom: 10px;
    }
    .step-number {
        background-color: #721f2f;
        color: #ffffff;
        font-weight: 800;
        font-size: 0.85rem;
        padding: 4px 10px;
        border-radius: 6px;
        letter-spacing: 0.5px;
    }
    .step-title {
        color: #111827;
        font-size: 1.2rem;
        font-weight: 700;
        margin: 0;
    }
    .step-instruction {
        color: #6b7280;
        font-size: 0.9rem;
        margin-top: -6px;
        margin-bottom: 16px;
    }

    /* Progress Indicator */
    .progress-bar-container {
        display: flex;
        justify-content: space-between;
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 12px 20px;
        margin-bottom: 24px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.03);
        flex-wrap: wrap;
        gap: 10px;
    }
    .progress-step {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 0.85rem;
        font-weight: 700;
        color: #6b7280;
    }
    .progress-step.active {
        color: #721f2f;
    }
    .progress-dot {
        width: 22px;
        height: 22px;
        border-radius: 50%;
        background-color: #e5e7eb;
        color: #4b5563;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
    }
    .progress-step.active .progress-dot {
        background-color: #721f2f;
        color: #ffffff;
    }

    /* Resource Cards */
    .resource-badge-ok {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-left: 5px solid #16a34a;
        color: #166534;
        padding: 12px 16px;
        border-radius: 8px;
        font-size: 0.9rem;
    }
    .resource-badge-err {
        background-color: #fef2f2;
        border: 1px solid #fecaca;
        border-left: 5px solid #dc2626;
        color: #991b1b;
        padding: 12px 16px;
        border-radius: 8px;
        font-size: 0.9rem;
    }

    /* Summary Box */
    .summary-box {
        background-color: #f9fafb;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 20px;
    }
    .summary-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
        gap: 12px;
    }
    .summary-item {
        font-size: 0.85rem;
    }
    .summary-label {
        color: #6b7280;
        font-weight: 600;
        display: block;
    }
    .summary-value {
        color: #111827;
        font-weight: 700;
        font-size: 0.95rem;
    }

    /* Styling for ALL Action Buttons & Download Buttons */
    div.stButton > button:first-child, 
    div.stDownloadButton > button:first-child {
        background-color: #721f2f !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        border: 1px solid #721f2f !important;
        border-radius: 8px !important;
        padding: 12px 20px !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 3px 8px rgba(114, 31, 47, 0.2);
    }
    div.stButton > button:first-child:hover,
    div.stDownloadButton > button:first-child:hover {
        background-color: #541622 !important;
        border-color: #541622 !important;
        transform: translateY(-1px);
        box-shadow: 0 5px 12px rgba(114, 31, 47, 0.3);
    }
    /* Ensure internal text inside download button is forced white */
    div.stDownloadButton > button:first-child p,
    div.stDownloadButton > button:first-child span {
        color: #ffffff !important;
        font-weight: 700 !important;
    }

    /* Custom Input Labels Override */
    .stSelectbox label {
        color: #374151 !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Dynamic HOD Logic
# ---------------------------------------------------------
def get_hod_name(subject_name):
    technical_keywords = [
        "woodwork", "woodworking", "welding", "upholstery", "maintenance", 
        "hairdressing", "hospitality", "office admin", "ict"
    ]
    clean_subj = str(subject_name).strip().lower()
    if any(tech in clean_subj for tech in technical_keywords):
        return "S. Van Schalkwyk"
    return "B. Koenze"

# ---------------------------------------------------------
# Subject Registry & File Mapping
# ---------------------------------------------------------
SUBJECTS_CONFIG = {
    "Afrikaans Home Language": {
        "pdf": "curriculums/afrikaans_hl.pdf",
        "icon": "📖",
        "prompt_hint": "Focus on Afrikaans language structures, reading comprehension, literature, and formal writing."
    },
    "Afrikaans First Additional Language": {
        "pdf": "curriculums/afrikaans_fal.pdf",
        "icon": "💬",
        "prompt_hint": "Focus on practical Afrikaans vocabulary, basic grammar, listening comprehension, and functional everyday conversation."
    },
    "English Home Language": {
        "pdf": "curriculums/english_hl.pdf",
        "icon": "📚",
        "prompt_hint": "Focus on English literary analysis, essay writing, advanced grammar, and critical language awareness."
    },
    "English First Additional Language": {
        "pdf": "curriculums/english_fal.pdf",
        "icon": "🗣️",
        "prompt_hint": "Focus on practical English vocabulary, sentence construction, transactional writing, and reading comprehension."
    },
    "Mathematics": {
        "pdf": "curriculums/mathematics.pdf",
        "icon": "📐",
        "prompt_hint": "Focus on practical mathematical calculations, spatial reasoning, measurements, and real-world vocational problem solving."
    },
    "Natural Science": {
        "pdf": "curriculums/natural_science.pdf",
        "icon": "🔬",
        "prompt_hint": "Focus on scientific observation, practical experiments, biological systems, physical processes, and environmental awareness."
    },
    "PSW": {
        "pdf": "curriculums/psw.pdf",
        "icon": "🌱",
        "prompt_hint": "Focus on personal development, emotional well-being, social responsibility, workplace ethics, and health education."
    },
    "Creative Arts": {
        "pdf": "curriculums/creative_arts.pdf",
        "icon": "🎨",
        "prompt_hint": "Focus on practical artistic techniques, visual arts, design concepts, drama, and creative expression."
    },
    "Physical Education": {
        "pdf": "curriculums/physical_education.pdf",
        "icon": "⚽",
        "prompt_hint": "Focus on motor skill development, fitness safety rules, practical sports execution, teamwork, and active participation."
    },
    "Hospitality": {
        "pdf": "curriculums/hospitality.pdf",
        "icon": "🍳",
        "prompt_hint": "Focus on commercial kitchen hygiene, food preparation techniques, culinary safety, customer service, and recipe execution."
    },
    "Office Admin": {
        "pdf": "curriculums/office_admin.pdf",
        "icon": "💻",
        "prompt_hint": "Focus on office filing systems, computer literacy, business documents, professional communication, and administrative workflow."
    },
    "Hairdressing": {
        "pdf": "curriculums/hairdressing.pdf",
        "icon": "✂️",
        "prompt_hint": "Focus on salon safety, scalp/hair treatments, styling tools, client consultation, and practical hairdressing techniques."
    },
    "Upholstery": {
        "pdf": "curriculums/upholstery.pdf",
        "icon": "🛋️",
        "prompt_hint": "Focus on fabric cutting, tool handling, furniture frame preparation, padding, stitching, and workshop safety."
    },
    "Woodwork": {
        "pdf": "curriculums/woodwork.pdf",
        "icon": "🪵",
        "prompt_hint": "Focus on workshop safety, timber selection, hand and power tool operation, joinery techniques, and project assembly."
    },
    "Welding": {
        "pdf": "curriculums/welding.pdf",
        "icon": "👨‍🏭",
        "prompt_hint": "Focus on PPE safety compliance, arc/MIG welding techniques, metal joint preparation, grinder operation, and structural strength."
    },
    "ICT": {
        "pdf": "curriculums/ict.pdf",
        "icon": "💻",
        "prompt_hint": "Focus on computer literacy, software applications, keyboarding, internet safety, digital communication, and basic hardware troubleshooting."
    },
    "Maintenance": {
        "pdf": "curriculums/maintenance.pdf",
        "icon": "🛠️",
        "prompt_hint": "Focus on general building maintenance, basic plumbing, electrical safety, painting, tool repair, and facility upkeep."
    }
}

GROQ_API_KEY = st.secrets.get("GROQ_API_KEY")

TEACHERS = [
    "K. Abrahams",
    "R. Adams",
    "E. Britz",
    "S. Cornelius",
    "D. Davids",
    "R. Desai",
    "U. De Villiers",
    "L. Grundlingh",
    "B. Koenze",
    "B. Manuel",
    "M. Murray",
    "A. Pressend",
    "C. September",
    "N. Smuts",
    "S. St. Jerry",
    "S. Van Schalkwyk",
    "A. Williams"
]

TERM_DATES_2026 = {
    1: {
        "1": "14 Jan – 16 Jan 2026",
        "2": "19 Jan – 23 Jan 2026",
        "3": "26 Jan – 30 Jan 2026",
        "4": "2 Feb – 6 Feb 2026",
        "5": "9 Feb – 13 Feb 2026",
        "6": "16 Feb – 20 Feb 2026",
        "7": "23 Feb – 27 Feb 2026",
        "8": "2 Mar – 6 Mar 2026",
        "9": "9 Mar – 13 Mar 2026",
        "10": "16 Mar – 20 Mar 2026",
        "11": "23 Mar – 27 Mar 2026",
        "1 – 2": "14 Jan – 23 Jan 2026",
        "3 – 4": "26 Jan – 6 Feb 2026",
        "5 – 6": "9 Feb – 20 Feb 2026",
        "7 – 8": "23 Feb – 6 Mar 2026",
        "9 – 10": "9 Mar – 20 Mar 2026",
        "10 – 11": "16 Mar – 27 Mar 2026",
    },
    2: {
        "1": "8 Apr – 10 Apr 2026",
        "2": "13 Apr – 17 Apr 2026",
        "3": "20 Apr – 24 Apr 2026",
        "4": "28 Apr – 30 Apr 2026",
        "5": "4 May – 8 May 2026",
        "6": "11 May – 15 May 2026",
        "7": "18 May – 22 May 2026",
        "8": "25 May – 29 May 2026",
        "9": "1 Jun – 5 Jun 2026",
        "10": "8 Jun – 12 Jun 2026",
        "11": "17 Jun – 19 Jun 2026",
        "12": "22 Jun – 26 Jun 2026",
        "1 – 2": "8 Apr – 17 Apr 2026",
        "3 – 4": "20 Apr – 30 Apr 2026",
        "5 – 6": "4 May – 15 May 2026",
        "7 – 8": "18 May – 29 May 2026",
        "9 – 10": "1 Jun – 12 Jun 2026",
        "11 – 12": "17 Jun – 26 Jun 2026",
    },
    3: {
        "1": "21 Jul – 24 Jul 2026",
        "2": "27 Jul – 31 Jul 2026",
        "3": "3 Aug – 7 Aug 2026",
        "4": "11 Aug – 14 Aug 2026",
        "5": "17 Aug – 21 Aug 2026",
        "6": "24 Aug – 28 Aug 2026",
        "7": "31 Aug – 4 Sep 2026",
        "8": "7 Sep – 11 Sep 2026",
        "9": "14 Sep – 18 Sep 2026",
        "10": "21 Sep – 23 Sep 2026",
        "1 – 2": "21 Jul – 31 Jul 2026",
        "1 – 3": "21 Jul – 7 Aug 2026",
        "2 – 3": "27 Jul – 7 Aug 2026",
        "3 – 4": "3 Aug – 14 Aug 2026",
        "4 – 5": "11 Aug – 21 Aug 2026",
        "5 – 6": "17 Aug – 28 Aug 2026",
        "6 – 7": "24 Aug – 4 Sep 2026",
        "7 – 8": "31 Aug – 11 Sep 2026",
        "8 – 9": "7 Sep – 18 Sep 2026",
        "9 – 10": "14 Sep – 23 Sep 2026",
        "8 – 10": "7 Sep – 23 Sep 2026",
    },
    4: {
        "1": "6 Oct – 9 Oct 2026",
        "2": "12 Oct – 16 Oct 2026",
        "3": "19 Oct – 23 Oct 2026",
        "4": "26 Oct – 30 Oct 2026",
        "5": "2 Nov – 6 Nov 2026",
        "6": "9 Nov – 13 Nov 2026",
        "7": "16 Nov – 20 Nov 2026",
        "8": "23 Nov – 27 Nov 2026",
        "9": "30 Nov – 4 Dec 2026",
        "10": "7 Dec – 9 Dec 2026",
        "1 – 2": "6 Oct – 16 Oct 2026",
        "3 – 4": "19 Oct – 30 Oct 2026",
        "5 – 6": "2 Nov – 13 Nov 2026",
        "7 – 8": "16 Nov – 27 Nov 2026",
        "9 – 10": "30 Nov – 9 Dec 2026",
    }
}

TERM_OVERALL_RANGES = {
    1: "14 Jan – 27 Mar 2026",
    2: "08 Apr – 26 Jun 2026",
    3: "21 Jul – 23 Sep 2026",
    4: "06 Oct – 09 Dec 2026"
}

# ---------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------
def get_badge_path():
    folder = os.path.dirname(os.path.abspath(__file__))
    for file in os.listdir(folder):
        if file.lower().endswith((".jpg", ".jpeg", ".png")) and "badge" in file.lower():
            return os.path.join(folder, file)
    fallback = os.path.join(folder, "florida_badge.jpg")
    return fallback if os.path.exists(fallback) else None

def get_base64_image(image_path):
    if image_path and os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode("utf-8")
    return None

def prevent_table_break(doc):
    for table in doc.tables:
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            trPr.append(OxmlElement('w:cantSplit'))
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    pPr = paragraph._p.get_or_add_pPr()
                    pPr.append(OxmlElement('w:keepNext'))

def get_automatic_date_range(term, week_str):
    term_dict = TERM_DATES_2026.get(term, {})
    clean_week = week_str.strip()
    if clean_week in term_dict:
        return term_dict[clean_week]
    return TERM_OVERALL_RANGES.get(term, "Term Date TBD")

def extract_pdf_section(pdf_path, year, term, subject_name=""):
    reader = PdfReader(pdf_path)
    year_words = {1: "ONE", 2: "TWO", 3: "THREE", 4: "FOUR"}
    
    y_patterns = [f"YEAR {year}", f"YEAR {year_words.get(year, '')}", f"GRADE {year}", f"STAGE {year}"]
    t_patterns = [f"TERM {term}"]
    
    matched_pages = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        txt_upper = txt.upper()
        
        has_year = any(p in txt_upper for p in y_patterns)
        has_term = any(p in txt_upper for p in t_patterns)
        
        if has_year and has_term:
            matched_pages.append(txt)
            
    if matched_pages:
        return "\n".join(matched_pages)
    
    year_only_pages = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        txt_upper = txt.upper()
        if any(p in txt_upper for p in y_patterns):
            year_only_pages.append(txt)
            
    if year_only_pages:
        return "\n".join(year_only_pages)

    return "\n".join([page.extract_text() or "" for page in reader.pages])

def format_to_str(val):
    if isinstance(val, list):
        items = []
        for item in val:
            if isinstance(item, dict):
                line = " - ".join([str(v) for v in item.values() if v])
                if line:
                    items.append(f"• {line}")
            else:
                items.append(f"• {str(item)}")
        return "\n".join(items)
    elif isinstance(val, dict):
        lines = []
        for v in val.values():
            if isinstance(v, list):
                lines.append(format_to_str(v))
            elif v:
                lines.append(f"• {str(v)}")
        return "\n".join(lines)
    elif val is None:
        return ""
    return str(val).strip()

class LessonPlanData(BaseModel):
    theme: str
    topics: str
    outcomes: str
    activities: str
    integration: str
    ltsm: str
    informal_assessment: str
    formal_assessment: str
    barriers: str
    skills: str

def extract_content(pdf_text, year, term, week, subject_name, prompt_hint=""):
    if not GROQ_API_KEY:
        raise ValueError("Groq API Key is missing. Check your secrets.toml file.")
        
    client = Groq(api_key=GROQ_API_KEY)
    
    prompt = f"""
    You are an expert curriculum developer for South African CAPS Technical Occupational schools.
    Generate a BALANCED, CLEAR, AND DETAILED single-page lesson plan for: {subject_name}.

    STRICT SUBJECT TARGETING:
    - Target Subject: {subject_name}.
    - Ignore any text in the PDF that belongs to other subjects or general introductions.

    SUBJECT FOCUS GUIDELINE:
    {prompt_hint}

    FIELD INSTRUCTIONS (BALANCED 1-PAGE QUALITY & DETAIL):
    - theme: Clean main unit topic/module title.
    - topics: 3 clear, practical focus bullet points for the week.
    - outcomes: 2-3 specific, measurable learning/practical outcomes.
    - activities: Detailed 3-stage breakdown:
      • Introduction: Brief hook, baseline check, or safety review.
      • Teacher Explanation/Demo: Direct instruction or practical demonstration.
      • Learner Practical Task: Hands-on exercise, workshop task, or practical application.
    - integration: 1-2 practical connections to workshop safety, tools, cross-subject skills, or real-world application.
    - ltsm: Detailed list of equipment, tools, PPE, workbooks, or physical materials needed.
    - informal_assessment: 2 practical observation points (e.g. checking work accuracy, inspecting tool handling).
    - formal_assessment: Formal task name or test focus, or "N/A - Informal observation week".
    - barriers: 2 actionable support strategies for struggling learners (e.g. visual step cards, peer pairing, extra practice time).
    - skills: 3-4 specific practical or technical skills developed.

    STYLE RULES:
    1. SIMPLE & DIRECT: Write clearly without filler words or overly academic jargon.
    2. NO NESTED KEYS: NEVER output keys like "description:", "skills:", or "guided_practical_examples:". Use simple bullet points (`•`).

    CURRICULUM SOURCE TEXT (Year {year}, Term {term}, Week {week}):
    {pdf_text[:15000]}

    Return ONLY raw JSON with clean string values for each key.
    """
    
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        response_format={"type": "json_object"}
    )
    
    raw_json = json.loads(response.choices[0].message.content)
    clean_json = {k: format_to_str(v) for k, v in raw_json.items()}
    
    return LessonPlanData.model_validate(clean_json)

# ---------------------------------------------------------
# 1. Compact Top Header
# ---------------------------------------------------------
badge_path = get_badge_path()
badge_b64 = get_base64_image(badge_path)
badge_html = f'<img src="data:image/jpeg;base64,{badge_b64}" class="header-badge-img"/>' if badge_b64 else ''

st.markdown(f"""
    <div class="compact-header">
        {badge_html}
        <div class="header-text-container">
            <h1>FLORIDA SCHOOL OF SKILLS</h1>
            <p>CAPS MULTI-SUBJECT LESSON PLAN GENERATOR</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 3. Progress Indicator
# ---------------------------------------------------------
st.markdown("""
    <div class="progress-bar-container">
        <div class="progress-step active"><div class="progress-dot">1</div> Subject & Teacher</div>
        <div class="progress-step active"><div class="progress-dot">2</div> Schedule</div>
        <div class="progress-step active"><div class="progress-dot">3</div> Resources</div>
        <div class="progress-step active"><div class="progress-dot">4</div> Generate</div>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 4. Step 1: Subject and Teacher
# ---------------------------------------------------------
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-number">STEP 1</span>
            <h3 class="step-title">Subject and Teacher</h3>
        </div>
        <p class="step-instruction">Select the subject and teacher for this lesson plan.</p>
""", unsafe_allow_html=True)

col_subj, col_teach = st.columns(2)
with col_subj:
    selected_subject = st.selectbox("Subject", list(SUBJECTS_CONFIG.keys()), key="step1_subject")
    subj_info = SUBJECTS_CONFIG[selected_subject]

with col_teach:
    teacher_name = st.selectbox("Teacher Name", TEACHERS, key="step1_teacher")

# Dynamic HOD Assignment
hod_name = get_hod_name(selected_subject)

st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 5. Step 2: Schedule Details
# ---------------------------------------------------------
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-number">STEP 2</span>
            <h3 class="step-title">Schedule Details</h3>
        </div>
        <p class="step-instruction">Select the year level, term, and week for the lesson plan.</p>
""", unsafe_allow_html=True)

col_yr, col_trm, col_wk = st.columns(3)
with col_yr:
    year = st.selectbox("Year Level", [1, 2, 3, 4], index=0, key="step2_year")

with col_trm:
    term = st.selectbox("Term", [1, 2, 3, 4], index=2, key="step2_term")

if term == 1:
    week_options = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "1 – 2", "3 – 4", "5 – 6", "7 – 8", "9 – 10", "10 – 11"]
elif term == 2:
    week_options = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "1 – 2", "3 – 4", "5 – 6", "7 – 8", "9 – 10", "11 – 12"]
elif term == 3:
    week_options = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "1 – 2", "1 – 3", "2 – 3", "3 – 4", "4 – 5", "5 – 6", "6 – 7", "7 – 8", "8 – 9", "9 – 10", "8 – 10"]
else:
    week_options = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "1 – 2", "3 – 4", "5 – 6", "7 – 8", "9 – 10"]

with col_wk:
    week_input = st.selectbox("Week(s)", week_options, key="step2_week")

auto_date_range = get_automatic_date_range(term, week_input)
st.caption(f"📅 **Auto Date Range:** {auto_date_range}")

st.markdown("</div>", unsafe_allow_html=True)

# Paths resolution
root_folder = os.path.dirname(os.path.abspath(__file__))
pdf_path = os.path.join(root_folder, subj_info["pdf"])
template_path = os.path.join(root_folder, "templates", "default_template.docx")

pdf_ready = os.path.exists(pdf_path)
template_ready = os.path.exists(template_path)
all_resources_ready = pdf_ready and template_ready

# ---------------------------------------------------------
# 6. Step 3: Resources and System Status
# ---------------------------------------------------------
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-number">STEP 3</span>
            <h3 class="step-title">Resources Ready</h3>
        </div>
""", unsafe_allow_html=True)

col_r1, col_r2 = st.columns(2)
with col_r1:
    if pdf_ready:
        st.markdown(f'''
            <div class="resource-badge-ok">
                <b>✓ Curriculum Loaded</b><br/>
                <small>{selected_subject} Curriculum ({os.path.basename(pdf_path)})</small>
            </div>
        ''', unsafe_allow_html=True)
    else:
        st.markdown(f'''
            <div class="resource-badge-err">
                <b>❌ Curriculum Missing</b><br/>
                <small>File not found: {subj_info["pdf"]}</small>
            </div>
        ''', unsafe_allow_html=True)

with col_r2:
    if template_ready:
        st.markdown('''
            <div class="resource-badge-ok">
                <b>✓ Word Template Loaded</b><br/>
                <small>default_template.docx</small>
            </div>
        ''', unsafe_allow_html=True)
    else:
        st.markdown('''
            <div class="resource-badge-err">
                <b>❌ Template Missing</b><br/>
                <small>File not found: templates/default_template.docx</small>
            </div>
        ''', unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 7. Step 4: Generate Lesson Plan
# ---------------------------------------------------------
st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <span class="step-number">STEP 4</span>
            <h3 class="step-title">Generate Lesson Plan</h3>
        </div>
""", unsafe_allow_html=True)

# Dynamic Summary Box (Includes HOD Display)
st.markdown(f"""
    <div class="summary-box">
        <div class="summary-grid">
            <div class="summary-item">
                <span class="summary-label">Subject</span>
                <span class="summary-value">{selected_subject}</span>
            </div>
            <div class="summary-item">
                <span class="summary-label">Teacher</span>
                <span class="summary-value">{teacher_name}</span>
            </div>
            <div class="summary-item">
                <span class="summary-label">Assigned HOD</span>
                <span class="summary-value">{hod_name}</span>
            </div>
            <div class="summary-item">
                <span class="summary-label">Year Level</span>
                <span class="summary-value">Year {year}</span>
            </div>
            <div class="summary-item">
                <span class="summary-label">Term & Week</span>
                <span class="summary-value">Term {term}, Week {week_input}</span>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# Button Text Logic
clean_subj_upper = selected_subject.upper()
if len(clean_subj_upper) > 22:
    btn_label = "GENERATE LESSON PLAN"
else:
    btn_label = f"GENERATE {clean_subj_upper} LESSON PLAN"

# Main Generation Action
if st.button(btn_label, disabled=not all_resources_ready):
    with st.spinner("Generating your lesson plan... Please wait while the document is prepared."):
        try:
            pdf_text = extract_pdf_section(pdf_path, year, term, selected_subject)
            extracted = extract_content(
                pdf_text=pdf_text,
                year=year,
                term=term,
                week=week_input,
                subject_name=selected_subject.upper(),
                prompt_hint=subj_info["prompt_hint"]
            )
            
            context = extracted.model_dump()
            context.update({
                "teacher_name": teacher_name,
                "hod": hod_name,
                "hod_name": hod_name,
                "year": str(year),
                "term": str(term),
                "week": str(week_input),
                "date_range": str(auto_date_range),
                "subject": selected_subject.upper()
            })

            doc = DocxTemplate(template_path)
            doc.render(context)
            prevent_table_break(doc.docx)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.session_state["generated_docx"] = buffer
            st.session_state["generated_filename"] = f"{selected_subject.replace(' ', '_')}_Lesson_Plan_Year{year}_Term{term}_Week{week_input}.docx"
            st.session_state["gen_details"] = {"subj": selected_subject, "yr": year, "trm": term, "wk": week_input, "hod": hod_name}

        except Exception as e:
            st.error(f"Error generating document: {e}")

# Success State Render
if "generated_docx" in st.session_state and st.session_state["generated_docx"] is not None:
    det = st.session_state["gen_details"]
    st.markdown(f"""
        <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-left: 5px solid #16a34a; padding: 20px; border-radius: 8px; margin-top: 20px; margin-bottom: 20px;">
            <h4 style="color: #166534; margin: 0 0 8px 0;">✓ Lesson Plan Generated Successfully</h4>
            <p style="color: #15803d; margin: 0 0 0 0; font-size: 0.95rem;">
                Your <b>{det['subj']}</b> lesson plan for <b>Year {det['yr']}, Term {det['trm']}, Week {det['wk']}</b> (HOD: <b>{det['hod']}</b>) is ready for download.
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    col_dl, col_rst = st.columns(2)
    with col_dl:
        st.download_button(
            label="DOWNLOAD LESSON PLAN",
            data=st.session_state["generated_docx"],
            file_name=st.session_state["generated_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with col_rst:
        if st.button("GENERATE ANOTHER LESSON PLAN"):
            del st.session_state["generated_docx"]
            st.rerun()

st.markdown("</div>", unsafe_allow_html=True)