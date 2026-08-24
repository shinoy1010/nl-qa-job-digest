import os
import json
import re
import smtplib
import time
import urllib.parse
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from apify_client import ApifyClient
from google import genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn

# ==========================================
# 1. CONFIGURATION & API KEYS
# ==========================================
def load_dotenv(path=".env"):
    """Minimal .env loader -- no extra dependency needed to stay fully
    self-contained. Reads KEY=VALUE lines into os.environ if not already set."""
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key, value = key.strip(), value.strip().strip('"').strip("'")
            os.environ.setdefault(key, value)


load_dotenv()

APIFY_API_TOKEN = os.environ["APIFY_API_TOKEN"]
APIFY_TASK_ID = os.environ["APIFY_TASK_ID"]
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]

# Email delivery is optional -- if these aren't set, the pipeline still runs
# and just skips sending the report by email (e.g. local runs without SMTP set up).
EMAIL_SENDER = os.environ.get("EMAIL_SENDER", "")
EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")
EMAIL_RECIPIENT = os.environ.get("EMAIL_RECIPIENT", "shinoy1010@gmail.com")

# Initialize Clients
apify_client = ApifyClient(APIFY_API_TOKEN)
gemini_client = genai.Client(api_key=GEMINI_API_KEY)

# Load your Master Resume
with open("master_resume.txt", "r", encoding="utf-8") as file:
    MASTER_RESUME = file.read()

# Create directory for tailored CVs
CV_DIR = "Tailored_CVs"
os.makedirs(CV_DIR, exist_ok=True)

# Create directory for tailored cover letters
CL_DIR = "Cover_Letters"
os.makedirs(CL_DIR, exist_ok=True)

# Create directory for HTML digest reports
REPORTS_DIR = "reports"
os.makedirs(REPORTS_DIR, exist_ok=True)


def safe_filename(company, target_role):
    """Shared naming so the CV, cover letter, and report links all agree.
    Strips anything that isn't filesystem-safe (e.g. '/' in "QA/QC Manager")."""
    safe_company = re.sub(r'[^a-zA-Z0-9]', '_', company)
    safe_role = re.sub(r'[^a-zA-Z0-9]', '_', target_role)
    return f"Shinoy_{safe_company}_{safe_role}"


# ==========================================
# 2. DOCX GENERATION FUNCTION
# ==========================================
# Compact, single-column layout mirroring the reference CV format:
# narrow margins, small point sizes, tight paragraph spacing, tab-aligned
# job title/dates. Designed to consistently land at 1-1.5 pages instead of
# the ~3 pages the old add_heading()-based layout produced.

NAVY = RGBColor(0x1B, 0x3A, 0x5C)
BLUE = RGBColor(0x1F, 0x5C, 0x96)
DARK = RGBColor(0x11, 0x11, 0x11)

SUMMARY_MAX_CHARS = 480       # ~3-4 sentences
BULLET_MAX_CHARS = 220
MAX_TAILORED_BULLETS = 4


def _clip(text, max_chars):
    text = (text or "").strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0].rstrip(",.;:") + "…"


def _set_compact_page(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)
    normal = doc.styles['Normal']
    normal.font.name = 'Calibri'
    normal.font.size = Pt(10)


def _para(doc, space_before=0, space_after=2):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    return p


def _run(paragraph, text, size=10, bold=False, color=None):
    run = paragraph.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return run


def _section_header(doc, text):
    p = _para(doc, space_before=8, space_after=3)
    _run(p, text, size=11, bold=True, color=NAVY)
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.makeelement(qn('w:pBdr'), {})
    bottom = pPr.makeelement(qn('w:bottom'), {
        qn('w:val'): 'single', qn('w:sz'): '4', qn('w:space'): '1', qn('w:color'): '1B3A5C'
    })
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def _label_line(doc, label, value, tab_at=1.7):
    p = _para(doc, space_after=1)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(tab_at))
    _run(p, label, size=9, bold=True, color=NAVY)
    _run(p, '\t' + value, size=9)


def _job_bullets(doc, title, dates, context, bullets):
    p = _para(doc, space_before=6, space_after=0)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(6.3), WD_TAB_ALIGNMENT.RIGHT)
    _run(p, title, size=10.5, bold=True, color=BLUE)
    _run(p, '\t' + dates, size=9, bold=True, color=BLUE)

    p = _para(doc, space_after=2)
    _run(p, context, size=8.5, color=DARK)

    for b in bullets:
        p = _para(doc, space_after=1)
        _run(p, '▸ ' + b, size=9)


def generate_cv(company, target_role, summary, bullets):
    doc = Document()
    _set_compact_page(doc)

    # Header / Contact Info
    p = _para(doc, space_after=0)
    _run(p, 'SHINOY TICKU', size=20, bold=True, color=NAVY)

    p = _para(doc, space_after=2)
    _run(p, target_role or 'QA Automation Engineer', size=12, bold=True, color=BLUE)

    p = _para(doc, space_after=6)
    _run(p, '06 3090 5555  |  shinoy1010@gmail.com  |  Amstelveen, NL  |  EU Residence Permit Holder', size=9)

    # Professional Summary (AI-tailored, length-capped so it can't push the doc to a new page)
    _section_header(doc, 'PROFESSIONAL SUMMARY')
    p = _para(doc, space_after=6)
    _run(p, _clip(summary, SUMMARY_MAX_CHARS), size=9.5)

    # Core Competencies
    _section_header(doc, 'CORE COMPETENCIES')
    _label_line(doc, 'DevOps & Cloud:', 'AWS · Docker · Kubernetes · Terraform · Ansible · Linux · Bash')
    _label_line(doc, 'CI/CD & Config:', 'Jenkins · Git')
    _label_line(doc, 'QA Automation:', 'Java Selenium · Playwright · BDD Cucumber · TDD · Python · TestNG')
    _label_line(doc, 'Database & API:', 'Oracle SQL · RestAssured API · Postman')
    _label_line(doc, 'Functional Skills:', 'Stakeholder Management · Requirement Analysis · Data Quality Assurance · '
                                            'Process Optimization · Release Management · Shift-Left Testing')

    # Professional Experience
    _section_header(doc, 'PROFESSIONAL EXPERIENCE')

    # Most recent role: AI-tailored bullets swapped in (capped in count/length)
    # instead of appended as a separate section -- keeps the JD-specific
    # tailoring without adding a page of extra content.
    tailored_bullets = [_clip(b.replace('•', '').replace('-', '').strip(), BULLET_MAX_CHARS)
                         for b in (bullets or [])][:MAX_TAILORED_BULLETS]
    tcs_bullets = tailored_bullets or [
        'SRE Collaboration & Defect Reduction: hosted targeted calls with SRE teams before deployment, reducing defect leakage.',
        'CI/CD and Cloud Orchestration: implemented CI/CD pipelines via Jenkins, provisioning test environments with Terraform and AWS.',
        'Framework Development: migrated a Selenium-Java framework to Playwright TypeScript, boosting execution efficiency by 20%.',
        'Backend Validation: executed complex Oracle SQL queries to validate massive datasets and ensure backend data integrity.',
    ]
    _job_bullets(
        doc, 'QA Automation Engineer  |  Tata Consultancy Services', 'Apr 2023 – Apr 2026',
        'Global Banking and Financial Institution | EU Client', tcs_bullets,
    )

    _job_bullets(
        doc, 'Test Analyst  |  Infosys Limited', 'Mar 2021 – Mar 2023',
        'Global Retail Banking | US Client', [
            'Greenfield Automation: built a Test Automation framework from scratch for payment processing systems using Java Selenium on BDD Cucumber.',
            'Containerization: containerized test environments using Docker and configured Puppet Infrastructure for QA.',
        ],
    )

    _job_bullets(
        doc, 'System Engineer  |  Infosys Limited', 'Oct 2017 – Feb 2021',
        'Global Healthcare (US Client) | E-Commerce (UK Client)', [
            'Workflow Automation: automated cross-browser testing with Selenium across e-commerce client websites.',
            'Process Optimization: documented 15+ SOPs, increasing team efficiency by 30%.',
        ],
    )

    # Education & Certifications
    _section_header(doc, 'EDUCATION & CERTIFICATIONS')
    _label_line(doc, 'Bachelor of Engineering:', 'Computer Science (HBO) · Chitkara University, India · 2013–2017', tab_at=2.0)
    _label_line(doc, 'ISTQB CTFL:', 'Foundation Level – Software Testing', tab_at=2.0)
    _label_line(doc, 'ISTQB CTFL-AT:', 'Agile Testing Extension', tab_at=2.0)
    _label_line(doc, 'Microsoft AZ-900:', 'Azure Fundamentals – Microsoft Certified', tab_at=2.0)

    # Recognition
    _section_header(doc, 'RECOGNITION')
    _label_line(doc, 'On-the-Spot Award – TCS (2024):', 'Exemplary efforts & major contribution in quality reporting', tab_at=2.6)
    _label_line(doc, 'Award of Recognition – Infosys (2019):', 'Outstanding contribution, building partnership with Ricoh', tab_at=2.6)
    _label_line(doc, 'Insta Award – Infosys (2018):', 'Improving service in ALS support area & service reporting', tab_at=2.6)

    # Save the file
    filename = f"{safe_filename(company, target_role)}.docx"
    doc.save(f"{CV_DIR}/{filename}")
    return filename


# ==========================================
# 2b. COVER LETTER DOCX GENERATION FUNCTION
# ==========================================
def generate_cover_letter(company, target_role, cover_letter_text):
    doc = Document()
    doc.add_paragraph('Shinoy Ticku')
    doc.add_paragraph('Amstelveen, Netherlands  |  06 3090 5555  |  shinoy1010@gmail.com  |  EU Residence Permit Holder')
    doc.add_paragraph(datetime.now().strftime('%B %d, %Y'))
    doc.add_paragraph(f'Re: {target_role} — {company}')
    doc.add_paragraph('')

    for para in cover_letter_text.split('\n'):
        if para.strip():
            doc.add_paragraph(para.strip())

    filename = f"{safe_filename(company, target_role)}_Cover_Letter.docx"
    doc.save(f"{CL_DIR}/{filename}")
    return filename


# ==========================================
# 3. FETCH JOBS FROM APIFY
# ==========================================
def fetch_latest_jobs():
    print("🚀 Triggering Apify Scraper...")
    run = apify_client.task(APIFY_TASK_ID).call()
    
    print("✅ Scraping complete. Fetching dataset...")
    return apify_client.dataset(run["defaultDatasetId"]).list_items().items

# ==========================================
# 4. GEMINI AI ANALYSIS & TAILORING
# ==========================================
def tailor_resume_for_job(job_title, company_name, job_description):
    prompt = f"""
    You are an expert technical recruiter.
    I am providing my Master Resume and a Job Description. 
    
    Master Resume:
    {MASTER_RESUME}
    
    Job Title: {job_title}
    Company: {company_name}
    Job Description:
    {job_description}
    
    CRITICAL COVER LETTER INSTRUCTION:
    1. Greeting: Scan the Job Description for the name of the recruiter or hiring manager.
       - If a specific name IS found, start the letter EXACTLY with: "Hi [Name]," (e.g., "Hi Sarah,").
       - If NO specific name is found, start the letter EXACTLY with: "Hi {company_name} Hiring Team," (e.g., "Hi Astek Hiring Team,").
       - Only fall back to the generic "Hi Hiring Team," if the company name itself is missing or unclear.
       - NEVER use the word "Dear".
       
    2. Mandatory Opening Paragraph: Leave a blank line after the greeting, and then output this EXACT paragraph, word-for-word:
    "Having recently relocated to the Netherlands, I am actively seeking a new opportunity. I hold a valid EU residence permit with full work authorization and am available to start immediately. As someone who values ownership of quality, I am very excited about the chance to make a visible impact within your team."
    
    After that exact opening, continue with 1-2 short paragraphs tailoring my skills to this specific role and mentioning the company name.
    
    TONE GUIDELINES: The tone MUST sound completely human-written, like a real person typed it in one sitting -- never like it came from an AI writing tool. Be humble, polite, and professional. Use simple, clear, everyday English. DO NOT use fancy vocabulary, corporate jargon, or overly enthusiastic AI buzzwords (e.g., completely avoid words like "thrilled", "delighted", "meticulous", "catalyst", "spearheaded", "ignite", "leverage", "seamless", "robust", "dynamic", "passionate", "elevate", "unlock", "game-changer", etc.). Keep it grounded, direct, and straightforward.

    AVOID THESE AI-WRITING TELLS specifically:
    - Don't use the "not just X, but Y" or "not only X, but also Y" construction.
    - Don't overuse em-dashes -- at most one per letter, and only if it reads naturally.
    - Don't write in perfectly balanced triads (e.g., "fast, reliable, and scalable") -- real people don't group everything in threes.
    - Vary sentence length naturally; don't make every sentence the same tidy medium length.
    - Don't open sentences with "Additionally," "Furthermore," "Moreover," or "In today's fast-paced environment."
    - Don't use rhetorical questions or motivational-poster phrasing.
    - It's fine to use a contraction or two (e.g., "I've", "I'm") if it reads naturally -- perfectly formal, contraction-free writing can itself read as AI-generated.
    - Keep it a little imperfect and specific rather than smooth and generic -- mention something concrete from the JD rather than vague enthusiasm.
    
    CLOSING GUIDELINES: DO NOT include generic, AI-sounding filler sentences at the end like "Thank you for your time and consideration," or "I look forward to discussing how I can support your goals." Just end the letter abruptly and professionally with:
    "Best regards,\nShinoy"
    
    SCORING GUARDRAIL 1 (resume keyword/skill match): Compare the job description directly against
    the skills, tools, and experience actually listed in my Master Resume above (e.g. Selenium,
    Playwright, Cypress, TestCafe, Postman, JMeter, API/web/mobile software testing, CI/CD,
    Agile/Scrum, ISTQB, test automation frameworks, programming/scripting for test scripts, manual
    testing, exploratory testing, test case design, sanity/smoke/regression testing, defect
    tracking/management, AI-assisted test generation, AI/LLM-powered test automation tools (e.g.
    Gemini Code Assist, GitHub Copilot, self-healing test automation, AI-driven test case
    generation, prompt-based test scripting, AI code review/quality tooling). Manual testing skill
    and AI-assisted testing skill are each a valid, first-class match on their own -- do not require
    overlap with traditional automation tools alone; a role centered on manual/exploratory QA, or on
    AI-augmented testing/QA tooling, still counts as a strong match. The Match Score must reflect
    genuine overlap with those specific resume keywords, not
    just a shared job title or the word "Test"/"QA"/"Engineer" appearing in the title. If the
    description's actual responsibilities and required skills do NOT substantially overlap with the
    Master Resume's listed skills -- for example it's a Product Manager, Business Analyst, Data
    Scientist, generic Software/AI Engineer role, or a hardware/physical-product test role (signals:
    production lines, manufacturing, PPE/safety gear, cleanrooms, calibration, electronics/PCB
    testing, machinery, field service/installation visits, an "Operations"/"Manufacturing" job
    category) -- cap the Match Score at 20 regardless of title or seniority overlap. Only score
    above 50 if the role's core responsibilities genuinely match the software testing/QA skills
    (manual, automated, or AI-assisted) on the Master Resume.

    SCORING GUARDRAIL 2 (language requirement): The candidate does not speak fluent Dutch and can
    only work in English-friendly roles. Scan the description for any requirement or preference for
    Dutch language proficiency (e.g. "Nederlandse taal", "Dutch fluency required", "vloeiend
    Nederlands", "beheerst de Nederlandse taal", "goede beheersing van het Nederlands", or the
    description being written entirely in Dutch with no indication that English is acceptable for
    day-to-day work). If Dutch fluency/proficiency is required or strongly preferred, cap the Match
    Score at 15 regardless of technical fit. Roles that only mention Dutch as "a plus" while stating
    English is the working language are NOT penalized.

    AI-SKILLS TAILORING INSTRUCTION: If the job description mentions AI in any testing/QA context
    (e.g. AI-assisted testing, AI-powered test automation, LLM-based tooling, Copilot/AI coding
    assistants, self-healing tests, AI in the SDLC), make sure the Tailored Summary, at least one
    Tailored Bullet, and the Cover Letter explicitly surface my real AI-assisted testing experience
    from the Master Resume -- specifically streamlining test generation and deployment scripts using
    Gemini Code Assist. Do not fabricate AI tools or experience beyond what is in the Master Resume;
    only phrase my existing Gemini Code Assist experience in a way that speaks to the JD's AI angle.
    If the JD has no AI angle, do not force it in.

    Respond STRICTLY with a JSON object using the following schema:
    {{
        "Match Score": <integer 0-100 based on realistic fit, respecting both scoring guardrails above>,
        "Target Role": "<Cleaned up version of the job title, max 5 words>",
        "Tailored Summary": "<A 3-4 sentence professional summary tailored to this specific company and role. Highlight my Automation, Playwright, or CI/CD experience if requested in the JD. If the JD has an AI-in-testing angle, also weave in my Gemini Code Assist / AI-assisted testing experience per the AI-SKILLS TAILORING INSTRUCTION.>",
        "Tailored Bullets": [
            "<Bullet 1 matching JD requirements>",
            "<Bullet 2 matching JD requirements>",
            "<Bullet 3 matching JD requirements>",
            "<Bullet 4 matching JD requirements -- if the JD has an AI-in-testing angle, make one of these bullets about my Gemini Code Assist / AI-assisted testing experience>"
        ],
        "Cover Letter": "<The complete cover letter following the exact greeting rules, starting with the mandatory paragraph, ending EXACTLY with 'Best regards,\\nShinoy', ensuring no cliché AI filler sentences are used, and mentioning my AI-assisted testing experience if the JD has an AI-in-testing angle per the AI-SKILLS TAILORING INSTRUCTION.>"
    }}
    """
    
    response = gemini_client.models.generate_content(
        model='gemini-3.5-flash-lite',
        contents=prompt,
        config={'response_mime_type': 'application/json'}
    )
    
    return json.loads(response.text)

REPORT_CSS = """
:root {
  --bg: #f2f5f5; --surface: #ffffff; --surface-2: #eaf0ef;
  --ink: #14201f; --ink-dim: #4c605d; --ink-faint: #7c8f8c;
  --line: #d7e1df; --accent: #0f6e5c;
  --good: #0f6e5c; --good-bg: #e2f2ee;
  --warn: #92600c; --warn-bg: #fbf0dd;
  --low: #6b6f76; --low-bg: #ececee;
  --link: #0b5c8f;
  --shadow: 0 1px 2px rgba(20,32,31,0.06), 0 6px 20px -8px rgba(20,32,31,0.18);
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #10201d; --surface: #16302b; --surface-2: #1c3934;
    --ink: #eaf3f1; --ink-dim: #a9c2bd; --ink-faint: #7f9994;
    --line: #2a4b44; --accent: #4fd6b5;
    --good: #4fd6b5; --good-bg: #1d3f38;
    --warn: #f0b754; --warn-bg: #3a2f16;
    --low: #93a29e; --low-bg: #22403a;
    --link: #7cc4ef;
    --shadow: 0 1px 2px rgba(0,0,0,0.3), 0 10px 30px -10px rgba(0,0,0,0.55);
  }
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--ink); font-family: "IBM Plex Sans", system-ui, sans-serif; padding: clamp(20px, 4vw, 56px); display: flex; justify-content: center; }
main { width: 100%; max-width: 1040px; }
header.page { display: flex; flex-direction: column; gap: 6px; margin-bottom: 22px; }
.eyebrow { font-family: "IBM Plex Mono", monospace; font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; color: var(--accent); font-weight: 600; }
h1 { font-size: clamp(24px, 3.2vw, 32px); line-height: 1.15; margin: 0; font-weight: 700; }
.sub { color: var(--ink-dim); font-size: 15px; max-width: 68ch; line-height: 1.5; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin: 22px 0 28px; }
.stat { background: var(--surface); border: 1px solid var(--line); border-radius: 10px; padding: 14px 16px; box-shadow: var(--shadow); }
.stat .n { font-family: "IBM Plex Mono", monospace; font-variant-numeric: tabular-nums; font-size: 24px; font-weight: 600; color: var(--accent); line-height: 1; }
.stat .l { margin-top: 6px; font-size: 12.5px; color: var(--ink-dim); }
.table-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 12px; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; min-width: 720px; }
thead th { text-align: left; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-faint); font-weight: 600; padding: 12px 14px; border-bottom: 1px solid var(--line); background: var(--surface-2); white-space: nowrap; }
tbody td { padding: 12px 14px; border-bottom: 1px solid var(--line); font-size: 14px; vertical-align: middle; }
tbody tr:last-child td { border-bottom: none; }
tbody tr:hover { background: var(--surface-2); }
.role { font-weight: 600; }
.company { color: var(--ink-dim); font-size: 12.5px; margin-top: 2px; }
.score { font-family: "IBM Plex Mono", monospace; font-variant-numeric: tabular-nums; font-weight: 600; font-size: 13px; padding: 3px 9px; border-radius: 999px; display: inline-block; white-space: nowrap; }
.score.good { background: var(--good-bg); color: var(--good); }
.score.warn { background: var(--warn-bg); color: var(--warn); }
.score.low { background: var(--low-bg); color: var(--low); }
a.joblink { color: var(--link); text-decoration: none; font-size: 13px; font-weight: 500; white-space: nowrap; }
a.joblink:hover { text-decoration: underline; }
.doc-btns { display: flex; gap: 6px; flex-wrap: nowrap; }
.doc-btn { display: inline-flex; align-items: center; font-size: 12.5px; font-weight: 600; padding: 6px 10px; border-radius: 7px; border: 1px solid var(--line); background: var(--surface-2); color: var(--ink); text-decoration: none; white-space: nowrap; }
.doc-btn:hover { border-color: var(--accent); color: var(--accent); }
footer.page { margin-top: 32px; padding-top: 16px; border-top: 1px solid var(--line); font-size: 12.5px; color: var(--ink-faint); line-height: 1.6; }
"""


def is_easy_apply(rec):
    """Handles both a native bool (fresh pipeline run) and a string like
    'True'/'False' (when dashboard_data is reloaded from a CSV, e.g. for
    local report/email regeneration or testing)."""
    return str(rec.get("Easy Apply", False)).strip().lower() in ("true", "1", "yes")


def generate_html_report(dashboard_data, current_date):
    """Builds a self-contained local HTML digest (same visual language as the
    NL QA Job Digest) with working links straight to the CV/cover-letter files
    on disk. No external hosting, no Claude usage -- pure local output."""
    sorted_data = sorted(dashboard_data, key=lambda r: r["Match Score"], reverse=True)
    total = len(sorted_data)
    strong = sum(1 for r in sorted_data if r["Match Score"] >= 50)
    top = sorted_data[0] if sorted_data else None

    def tier(score):
        if score >= 50:
            return "good"
        if score >= 30:
            return "warn"
        return "low"

    row_tpl = """<tr>
      <td><span class="score {tier}">{score}</span></td>
      <td><div class="role">{role}</div><div class="company">{company}</div></td>
      <td><a class="joblink" href="{apply_link}" target="_blank" rel="noopener" title="{easy_apply_title}">{easy_apply_icon}View posting</a></td>
      <td><div class="doc-btns">
        <a class="doc-btn" href="{cv_href}">Resume</a>
        <a class="doc-btn" href="{cl_href}">Cover letter</a>
      </div></td>
    </tr>"""

    raw_base = "https://github.com/shinoy1010/job-automation-engine/raw/main"
    rows_html = []
    for rec in sorted_data:
        cv_file = rec.get("CV File", "")
        cl_file = rec.get("CL File", "")
        cv_href = f"{raw_base}/{CV_DIR}/{urllib.parse.quote(cv_file)}" if cv_file else "#"
        cl_href = f"{raw_base}/{CL_DIR}/{urllib.parse.quote(cl_file)}" if cl_file else "#"
        easy = is_easy_apply(rec)
        rows_html.append(row_tpl.format(
            tier=tier(rec["Match Score"]),
            score=rec["Match Score"],
            role=rec["Target Role"],
            company=rec["Company"],
            apply_link=rec.get("Apply Link", "") or "#",
            easy_apply_icon="⚡ " if easy else "",
            easy_apply_title="Easy Apply" if easy else "",
            cv_href=cv_href,
            cl_href=cl_href,
        ))

    html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NL QA Job Digest — {current_date}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<style>{REPORT_CSS}</style>
</head>
<body>
<main>
  <header class="page">
    <div class="eyebrow">NL QA / Test Automation Scan &middot; {current_date}</div>
    <h1>{total} roles evaluated, {strong} strong matches</h1>
    <p class="sub">Sourced via Apify and scored/tailored with Gemini against Shinoy's master resume. Every row links to a ready-to-review CV and cover letter on disk. ⚡ next to a posting means it supports LinkedIn Easy Apply.</p>
  </header>
  <div class="stats">
    <div class="stat"><div class="n">{total}</div><div class="l">jobs evaluated</div></div>
    <div class="stat"><div class="n">{strong}</div><div class="l">score 50+</div></div>
    <div class="stat"><div class="n">{top['Match Score'] if top else 0}</div><div class="l">top score &mdash; {top['Company'] if top else '—'}</div></div>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>Score</th><th>Role</th><th>Job link</th><th>Documents</th></tr></thead>
      <tbody>{''.join(rows_html)}</tbody>
    </table>
  </div>
  <footer class="page">Generated locally by app_engine.py &middot; CVs in Tailored_CVs/ &middot; cover letters in Cover_Letters/ &middot; nothing has been submitted anywhere.</footer>
</main>
</body>
</html>"""

    report_path = f"{REPORTS_DIR}/QA_Job_Digest_{current_date}.html"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html)
    return report_path


def generate_email_html(dashboard_data, current_date):
    """Builds an email-client-safe version of the digest: table-based layout,
    fully inline styles, no CSS grid/flexbox/variables and no external
    stylesheet link -- Gmail and other clients strip or mangle all of those,
    which is why the regular generate_html_report() output looks broken when
    emailed directly."""
    sorted_data = sorted(dashboard_data, key=lambda r: r["Match Score"], reverse=True)
    total = len(sorted_data)
    strong = sum(1 for r in sorted_data if r["Match Score"] >= 50)
    top = sorted_data[0] if sorted_data else None

    def tier_color(score):
        if score >= 50:
            return "#0f6e5c", "#e2f2ee"   # good: text, bg
        if score >= 30:
            return "#92600c", "#fbf0dd"   # warn
        return "#6b6f76", "#ececee"       # low

    raw_base = "https://github.com/shinoy1010/job-automation-engine/raw/main"
    font = "font-family: Arial, Helvetica, sans-serif;"

    def stat_cell(number, label):
        return f"""<td width="33%" valign="top" style="padding:10px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #d7e1df;border-radius:8px;">
            <tr><td style="padding:12px 14px;">
              <div style="{font}font-size:22px;font-weight:700;color:#0f6e5c;">{number}</div>
              <div style="{font}font-size:12px;color:#4c605d;margin-top:4px;">{label}</div>
            </td></tr>
          </table>
        </td>"""

    row_rows = []
    for rec in sorted_data:
        score = rec["Match Score"]
        text_color, bg_color = tier_color(score)
        cv_file = rec.get("CV File", "")
        cl_file = rec.get("CL File", "")
        cv_href = f"{raw_base}/{CV_DIR}/{urllib.parse.quote(cv_file)}" if cv_file else "#"
        cl_href = f"{raw_base}/{CL_DIR}/{urllib.parse.quote(cl_file)}" if cl_file else "#"
        apply_link = rec.get("Apply Link", "") or "#"
        easy_apply_icon = "⚡ " if is_easy_apply(rec) else ""
        row_rows.append(f"""
        <tr>
          <td style="{font}padding:10px 8px;border-bottom:1px solid #e5e9e8;vertical-align:top;">
            <span style="display:inline-block;padding:3px 8px;border-radius:6px;font-size:12px;font-weight:700;color:{text_color};background:{bg_color};">{score}</span>
          </td>
          <td style="{font}padding:10px 8px;border-bottom:1px solid #e5e9e8;vertical-align:top;">
            <div style="font-size:13.5px;font-weight:600;color:#14201f;">{rec['Target Role']}</div>
            <div style="font-size:12.5px;color:#4c605d;margin-top:2px;">{rec['Company']}</div>
          </td>
          <td style="{font}padding:10px 8px;border-bottom:1px solid #e5e9e8;vertical-align:top;white-space:nowrap;">
            <a href="{apply_link}" title="{'Easy Apply' if easy_apply_icon else ''}" style="font-size:12.5px;color:#0b5c8f;text-decoration:none;">{easy_apply_icon}View posting</a>
          </td>
          <td style="{font}padding:10px 8px;border-bottom:1px solid #e5e9e8;vertical-align:top;white-space:nowrap;">
            <a href="{cv_href}" style="display:inline-block;font-size:12px;font-weight:600;padding:5px 9px;margin-right:4px;border:1px solid #d7e1df;border-radius:6px;background:#eaf0ef;color:#14201f;text-decoration:none;">Resume</a>
            <a href="{cl_href}" style="display:inline-block;font-size:12px;font-weight:600;padding:5px 9px;border:1px solid #d7e1df;border-radius:6px;background:#eaf0ef;color:#14201f;text-decoration:none;">Cover letter</a>
          </td>
        </tr>""")

    html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NL QA Job Digest — {current_date}</title>
</head>
<body style="margin:0;padding:0;background:#f2f5f5;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f2f5f5;">
<tr><td align="center" style="padding:24px 12px;">
<table role="presentation" width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;background:#ffffff;border:1px solid #d7e1df;border-radius:10px;">
  <tr><td style="padding:24px 24px 4px 24px;">
    <div style="{font}font-size:11px;letter-spacing:0.08em;text-transform:uppercase;color:#0f6e5c;font-weight:700;">NL QA / TEST AUTOMATION SCAN &middot; {current_date}</div>
    <div style="{font}font-size:22px;font-weight:700;color:#14201f;margin-top:6px;line-height:1.25;">{total} roles evaluated, {strong} strong matches</div>
    <div style="{font}font-size:13.5px;color:#4c605d;margin-top:8px;line-height:1.5;">Sourced via Apify and scored/tailored with Gemini against Shinoy's master resume. Every row links to a ready-to-review CV and cover letter. ⚡ next to a posting means it supports LinkedIn Easy Apply.</div>
  </td></tr>
  <tr><td style="padding:14px 16px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
      {stat_cell(total, "jobs evaluated")}
      {stat_cell(strong, "score 50+")}
      {stat_cell(f"{top['Match Score']}" if top else 0, f"top score &mdash; {top['Company']}" if top else "—")}
    </tr></table>
  </td></tr>
  <tr><td style="padding:4px 24px 8px 24px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
      <tr>
        <th align="left" style="{font}font-size:11px;text-transform:uppercase;color:#4c605d;padding:8px;border-bottom:2px solid #d7e1df;">Score</th>
        <th align="left" style="{font}font-size:11px;text-transform:uppercase;color:#4c605d;padding:8px;border-bottom:2px solid #d7e1df;">Role</th>
        <th align="left" style="{font}font-size:11px;text-transform:uppercase;color:#4c605d;padding:8px;border-bottom:2px solid #d7e1df;">Job link</th>
        <th align="left" style="{font}font-size:11px;text-transform:uppercase;color:#4c605d;padding:8px;border-bottom:2px solid #d7e1df;">Documents</th>
      </tr>
      {''.join(row_rows)}
    </table>
  </td></tr>
  <tr><td style="padding:16px 24px 24px 24px;border-top:1px solid #eaf0ef;">
    <div style="{font}font-size:12px;color:#7c8f8c;line-height:1.6;">Generated by app_engine.py &middot; nothing has been submitted anywhere -- every document is a draft for manual review. View the full report on <a href="https://github.com/shinoy1010/job-automation-engine" style="color:#0b5c8f;">GitHub</a>.</div>
  </td></tr>
</table>
</td></tr>
</table>
</body>
</html>"""
    return html


def generate_markdown_report(dashboard_data, current_date):
    """GitHub renders .md files inline in the repo browser (unlike .html,
    which it shows as raw source), so this is the version to open directly
    on github.com -- no local machine, no Pages, repo stays private."""
    sorted_data = sorted(dashboard_data, key=lambda r: r["Match Score"], reverse=True)
    total = len(sorted_data)
    strong = sum(1 for r in sorted_data if r["Match Score"] >= 50)
    top = sorted_data[0] if sorted_data else None

    def md_escape(text):
        return (text or "").replace("|", "\\|").replace("\n", " ")

    top_line = ""
    if top:
        top_line = f"Top pick: **{top['Company']}** — {top['Target Role']} (score {top['Match Score']})."

    lines = [
        f"# NL QA Job Digest — {current_date}",
        "",
        f"**{total} roles evaluated, {strong} scored 50+.** {top_line}",
        "",
        "Sourced via Apify, scored and tailored with Gemini against the master resume. "
        "Nothing has been submitted anywhere -- every document is a draft for manual review. "
        "⚡ next to a posting means it supports LinkedIn Easy Apply.",
        "",
        "| Score | Role | Company | Job link | Resume | Cover letter |",
        "|---|---|---|---|---|---|",
    ]

    # Plain relative links resolve to GitHub's "blob" preview page (an extra
    # click to actually download). The /raw/ path serves the file directly,
    # so clicking Resume/Cover letter downloads it immediately instead.
    raw_base = "https://github.com/shinoy1010/job-automation-engine/raw/main"

    for rec in sorted_data:
        cv_file = rec.get("CV File", "")
        cl_file = rec.get("CL File", "")
        cv_link = f"[Resume]({raw_base}/{CV_DIR}/{urllib.parse.quote(cv_file)})" if cv_file else "—"
        cl_link = f"[Cover letter]({raw_base}/{CL_DIR}/{urllib.parse.quote(cl_file)})" if cl_file else "—"
        apply_link = rec.get("Apply Link", "") or "#"
        easy_apply_icon = "⚡ " if is_easy_apply(rec) else ""
        lines.append(
            f"| {rec['Match Score']} | {md_escape(rec['Target Role'])} | {md_escape(rec['Company'])} "
            f"| {easy_apply_icon}[View posting]({apply_link}) | {cv_link} | {cl_link} |"
        )

    report_path = f"{REPORTS_DIR}/QA_Job_Digest_{current_date}.md"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return report_path


def send_email_report(dashboard_data, current_date):
    """Emails the NL QA Job Digest via Gmail SMTP once the pipeline finishes
    successfully, using the email-safe inline-styled template (see
    generate_email_html) rather than the CSS-grid/variables version built for
    browsers/GitHub, which Gmail renders broken. Silently skipped (with a
    console warning) if EMAIL_SENDER/EMAIL_APP_PASSWORD aren't configured --
    keeps this optional and self-contained, no external mail API or extra
    dependency needed."""
    if not EMAIL_SENDER or not EMAIL_APP_PASSWORD:
        print("   ⚠️  Email not sent -- EMAIL_SENDER/EMAIL_APP_PASSWORD not configured.")
        return

    job_count = len(dashboard_data)
    html_content = generate_email_html(dashboard_data, current_date)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"NL QA Job Digest — {current_date} ({job_count} matches)"
    msg["From"] = EMAIL_SENDER
    msg["To"] = EMAIL_RECIPIENT
    msg.attach(MIMEText(html_content, "html"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_APP_PASSWORD)
            server.sendmail(EMAIL_SENDER, [EMAIL_RECIPIENT], msg.as_string())
        print(f"   📧 Report emailed to {EMAIL_RECIPIENT}")
    except Exception as e:
        print(f"   ❌ Failed to email report: {e}")


# ==========================================
# 5. MAIN EXECUTION PIPELINE
# ==========================================
def run_pipeline():
    jobs = fetch_latest_jobs()
    dashboard_data = []
    total_jobs = len(jobs)

    print(f"🔍 Found {total_jobs} jobs. Filtering and evaluating all listings...")

    for index, job in enumerate(jobs):
        title = job.get("positionName", job.get("title", job.get("jobTitle", "Unknown Title")))
        company = job.get("companyName", job.get("company", "Unknown Company"))
        url = job.get("jobUrl", job.get("url", ""))
        description = job.get("description", job.get("jobDescription", job.get("text", "")))
        is_easy_apply = job.get("applyType") == "EASY_APPLY"

        if not description:
            continue
            
        desc_lower = description.lower()
        mentions_dutch = re.search(r'\b(dutch|nederlands|nederlandse)\b', desc_lower)
        is_optional = (
            re.search(r'\b(dutch|nederlands|nederlandse)\b.{0,60}?(plus|advantage|bonus|nice to have|pre|optional|not required)\b', desc_lower) or 
            re.search(r'\b(plus|advantage|bonus|nice to have|pre|optional|not required)\b.{0,60}?\b(dutch|nederlands|nederlandse)\b', desc_lower)
        )
                      
        if mentions_dutch and not is_optional:
            print(f"[{index + 1}/{total_jobs}] 🇳🇱 Skipping {company} - Dutch language strictly required.")
            continue
            
        print(f"[{index + 1}/{total_jobs}] Analyzing: {title} @ {company}")
        
        try:
            ai_data = tailor_resume_for_job(title, company, description)
            
            match_score = int(ai_data.get("Match Score", 0))
            if match_score < 40:
                print(f"   📉 Skipping {company} - Match score too low ({match_score}/100)")
                time.sleep(6) 
                continue

            # --- GENERATE WORD DOCUMENTS: CV + COVER LETTER ---
            target_role = ai_data.get("Target Role", title)
            cover_letter_text = ai_data.get("Cover Letter", "")

            cv_filename = generate_cv(
                company=company,
                target_role=target_role,
                summary=ai_data.get("Tailored Summary", ""),
                bullets=ai_data.get("Tailored Bullets", [])
            )
            cl_filename = generate_cover_letter(
                company=company,
                target_role=target_role,
                cover_letter_text=cover_letter_text
            )
            # ---------------------------------------------------

            formatted_bullets = "\n".join([f"• {bullet}" for bullet in ai_data["Tailored Bullets"]])

            record = {
                "Match Score": match_score,
                "Target Role": target_role,
                "Tailored Summary": ai_data.get("Tailored Summary", ""),
                "Tailored Bullets": formatted_bullets,
                "Company": company,
                "Job Description": description,
                "Apply Link": url,
                "Easy Apply": is_easy_apply,
                "Cover Letter": cover_letter_text,
                "CV File": cv_filename,
                "CL File": cl_filename
            }
            dashboard_data.append(record)
            
            print(f"   ✅ Saved {company} Data & Generated CV (Score: {match_score})")
            time.sleep(6)
            
        except Exception as e:
            print(f"❌ Error processing {company}: {e}")
            time.sleep(10)

    if dashboard_data:
        df = pd.DataFrame(dashboard_data)
        df = df.sort_values(by="Match Score", ascending=False)
        
        now = datetime.now()
        current_date = f"{now.day}-{now.strftime('%b-%Y')}"
        csv_filename = f"QA_roles_{current_date}.csv"
        
        df.to_csv(csv_filename, index=False, encoding="utf-8")

        html_report_path = generate_html_report(dashboard_data, current_date)
        md_report_path = generate_markdown_report(dashboard_data, current_date)

        send_email_report(dashboard_data, current_date)

        print(f"\n🎉 Success! Evaluated all listings. {len(dashboard_data)} high-match jobs saved to '{csv_filename}'.")
        print(f"📄 Tailored Word documents are waiting for you in the 'Tailored_CVs' folder!")
        print(f"✉️  Tailored cover letters are waiting for you in the 'Cover_Letters' folder!")
        print(f"🌐 Local HTML digest: {html_report_path} (open it in your browser)")
        print(f"📋 GitHub-viewable digest: {md_report_path} (renders inline once pushed -- no local machine needed)")
    else:
        print("\n⚠️ No jobs with a match score above 40 were found.")

if __name__ == "__main__":
    run_pipeline()
