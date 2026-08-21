import os
import json
import re
import time
import urllib.parse
from datetime import datetime
import pandas as pd
from apify_client import ApifyClient
from google import genai
from docx import Document

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
    """Shared naming so the CV, cover letter, and report links all agree."""
    safe_company = re.sub(r'[^a-zA-Z0-9]', '_', company)
    safe_role = target_role.replace(' ', '_')
    return f"Shinoy_{safe_company}_{safe_role}"


# ==========================================
# 2. DOCX GENERATION FUNCTION
# ==========================================
def generate_cv(company, target_role, summary, bullets):
    doc = Document()
    
    # Header / Contact Info
    doc.add_heading('SHINOY TICKU', 0)
    doc.add_paragraph('QA Automation Engineer').bold = True
    doc.add_paragraph('06 3090 5555  |  shinoy1010@gmail.com  |  Amstelveen, NL  |  EU Residence Permit Holder')
    
    # Professional Summary (AI Tailored)
    doc.add_heading(' PROFESSIONAL SUMMARY', level=1)
    doc.add_paragraph(summary)
    
    # Highlighted Qualifications (AI Tailored)
    doc.add_heading(' KEY QUALIFICATIONS FOR TARGET ROLE', level=1)
    for bullet in bullets:
        clean_bullet = bullet.replace('•', '').replace('-', '').strip()
        doc.add_paragraph(f'▸ {clean_bullet}')
        
    # Core Competencies
    doc.add_heading(' CORE COMPETENCIES', level=1)
    doc.add_paragraph('DevOps & Cloud:\tAWS · Docker · Kubernetes · Terraform · Ansible · Linux · Bash')
    doc.add_paragraph('CI/CD & Config:\tJenkins · Git')
    doc.add_paragraph('QA Automation:\tJava Selenium · Playwright · BDD Cucumber · TDD · Python · TestNG')
    doc.add_paragraph('Database & API:\tOracle SQL · RestAssured API · Postman')
    doc.add_paragraph('Functional Skills:\tStakeholder Management · Requirement Analysis · Data Quality Assurance · Process Optimization · Release Management · Shift-Left Testing')

    # Professional Experience
    doc.add_heading(' PROFESSIONAL EXPERIENCE', level=1)
    
    # TCS Experience
    p_tcs = doc.add_paragraph()
    p_tcs.add_run('QA Automation Engineer  |  Tata Consultancy Services').bold = True
    p_tcs.add_run('\tApr 2023 – Apr 2026')
    doc.add_paragraph('Global Banking and Financial Institution | EU Client')
    doc.add_paragraph('▸ SRE Collaboration & Defect Reduction: scheduled and hosted targeted calls with SRE teams to explain release objectives and user stories prior to deployment, successfully identifying critical bottleneck areas which directly reduced defect leakage.')
    doc.add_paragraph('▸ CI/CD and Cloud Orchestration: implemented CI/CD pipelines and ran scheduled automation regression test builds via Jenkins while provisioning underlying test environments using Terraform and AWS to maintain high availability.')
    doc.add_paragraph('▸ Framework Development: migrated a Selenium-Java framework to Playwright TypeScript for credit limit and rating analysis end-to-end workflows within a microservices architecture, boosting execution efficiency by 20%.')
    doc.add_paragraph('▸ Backend Validation: executed complex Oracle SQL queries to validate massive datasets and ensure strict data integrity across backend credit risk workflows.')
    doc.add_paragraph('▸ Test Coverage: integrated Sealights to fetch comprehensive test coverage reports and utilized Allure reports to visualize test health.')
    doc.add_paragraph('▸ AI Tool Deployment: streamlined complex test generation and deployment scripts utilizing AI tools like Gemini Code Assist.')
    
    # Infosys Experience 1
    p_inf1 = doc.add_paragraph()
    p_inf1.add_run('Test Analyst  |  Infosys Limited').bold = True
    p_inf1.add_run('\tMar 2021 – Mar 2023')
    doc.add_paragraph('Global Retail Banking | US Client')
    doc.add_paragraph('▸ Greenfield Automation: created a Test Automation framework from scratch to test payment processing systems (CitiFT, GPP, IP, Isochaps, Volpay) utilizing Java Selenium on a BDD Cucumber Framework.')
    doc.add_paragraph('▸ Containerization and Config: containerized test environments using Docker and implemented configured Puppet Infrastructure for the QA environment.')
    doc.add_paragraph('▸ Continuous Validation: performed comprehensive exploratory testing alongside automated smoke testing and sanity runs on deployment builds to ensure rapid feedback cycles and high release stability.')
    doc.add_paragraph('▸ Defect Management: managed defect tracking on JIRA and maintained HP ALM with updated information to ensure alignment with business requirements.')

    # Infosys Experience 2
    p_inf2 = doc.add_paragraph()
    p_inf2.add_run('System Engineer  |  Infosys Limited').bold = True
    p_inf2.add_run('\tOct 2017 – Feb 2021')
    doc.add_paragraph('Global Healthcare (US Client) | E-Commerce (UK Client)')
    doc.add_paragraph('▸ Workflow Automation: automated redundant operational tasks using Selenium to execute comprehensive cross-browser testing across various e-commerce client websites.')
    doc.add_paragraph('▸ Process Optimization: documented over 15 Standard Operating Procedures (SOPs), successfully increasing overall team efficiency by 30%.')
    doc.add_paragraph('▸ Service Reporting: contributed directly to service reporting (WSR), incident management, and analyzed ServiceNow dumps for the PMO team.')
    
    # Education & Certifications
    doc.add_heading(' EDUCATION & CERTIFICATIONS', level=1)
    doc.add_paragraph('Bachelor of Engineering:\tComputer Science (HBO) · Chitkara University, India · 2013–2017')
    doc.add_paragraph('ISTQB CTFL:\tFoundation Level – Software Testing')
    doc.add_paragraph('ISTQB CTFL-AT:\tAgile Testing Extension')
    doc.add_paragraph('Microsoft AZ-900:\tAzure Fundamentals – Microsoft Certified')

    # Recognition
    doc.add_heading(' RECOGNITION', level=1)
    doc.add_paragraph('On-the-Spot Award – TCS (2024):\tExemplary efforts & major contribution in quality reporting')
    doc.add_paragraph('Award of Recognition – Infosys (2019):\tOutstanding contribution, building partnership with Ricoh')
    doc.add_paragraph('Insta Award – Infosys (2018):\tImproving service in ALS support area & service reporting')
    
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
       - If NO specific name is found, start the letter EXACTLY with: "Hi Hiring Team,"
       - NEVER use the word "Dear".
       
    2. Mandatory Opening Paragraph: Leave a blank line after the greeting, and then output this EXACT paragraph, word-for-word:
    "Having recently relocated to the Netherlands, I am actively seeking a new opportunity. I hold a valid EU residence permit with full work authorization and am available to start immediately. As someone who values ownership of quality, I am very excited about the chance to make a visible impact within your team."
    
    After that exact opening, continue with 1-2 short paragraphs tailoring my skills to this specific role and mentioning the company name.
    
    TONE GUIDELINES: The tone MUST sound completely human-written. Be humble, polite, and professional. Use simple, clear, everyday English. DO NOT use fancy vocabulary, corporate jargon, or overly enthusiastic AI buzzwords (e.g., completely avoid words like "thrilled", "delighted", "meticulous", "catalyst", "spearheaded", "ignite", etc.). Keep it grounded, direct, and straightforward.
    
    CLOSING GUIDELINES: DO NOT include generic, AI-sounding filler sentences at the end like "Thank you for your time and consideration," or "I look forward to discussing how I can support your goals." Just end the letter abruptly and professionally with:
    "Best regards,\nShinoy"
    
    Respond STRICTLY with a JSON object using the following schema:
    {{
        "Match Score": <integer 0-100 based on realistic fit>,
        "Target Role": "<Cleaned up version of the job title, max 5 words>",
        "Tailored Summary": "<A 3-4 sentence professional summary tailored to this specific company and role. Highlight my Automation, Playwright, or CI/CD experience if requested in the JD.>",
        "Tailored Bullets": [
            "<Bullet 1 matching JD requirements>",
            "<Bullet 2 matching JD requirements>",
            "<Bullet 3 matching JD requirements>",
            "<Bullet 4 matching JD requirements>"
        ],
        "Cover Letter": "<The complete cover letter following the exact greeting rules, starting with the mandatory paragraph, ending EXACTLY with 'Best regards,\\nShinoy', and ensuring no cliché AI filler sentences are used.>"
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


def generate_html_report(dashboard_data, current_date):
    """Builds a self-contained local HTML digest (same visual language as the
    NL QA Job Digest) with working links straight to the CV/cover-letter files
    on disk. No external hosting, no Claude usage -- pure local output."""
    sorted_data = sorted(dashboard_data, key=lambda r: r["Match Score"], reverse=True)
    total = len(sorted_data)
    strong = sum(1 for r in sorted_data if r["Match Score"] >= 70)
    top = sorted_data[0] if sorted_data else None

    def tier(score):
        if score >= 70:
            return "good"
        if score >= 50:
            return "warn"
        return "low"

    row_tpl = """<tr>
      <td><span class="score {tier}">{score}</span></td>
      <td><div class="role">{role}</div><div class="company">{company}</div></td>
      <td><a class="joblink" href="{apply_link}" target="_blank" rel="noopener">View posting</a></td>
      <td><div class="doc-btns">
        <a class="doc-btn" href="{cv_href}">Resume</a>
        <a class="doc-btn" href="{cl_href}">Cover letter</a>
      </div></td>
    </tr>"""

    rows_html = []
    for rec in sorted_data:
        cv_href = urllib.parse.quote(f"../{CV_DIR}/{rec.get('CV File', '')}")
        cl_href = urllib.parse.quote(f"../{CL_DIR}/{rec.get('CL File', '')}")
        rows_html.append(row_tpl.format(
            tier=tier(rec["Match Score"]),
            score=rec["Match Score"],
            role=rec["Target Role"],
            company=rec["Company"],
            apply_link=rec.get("Apply Link", "") or "#",
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
    <p class="sub">Sourced via Apify and scored/tailored with Gemini against Shinoy's master resume. Every row links to a ready-to-review CV and cover letter on disk.</p>
  </header>
  <div class="stats">
    <div class="stat"><div class="n">{total}</div><div class="l">jobs evaluated</div></div>
    <div class="stat"><div class="n">{strong}</div><div class="l">score 70+</div></div>
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
            if match_score <= 29:
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

        report_path = generate_html_report(dashboard_data, current_date)

        print(f"\n🎉 Success! Evaluated all listings. {len(dashboard_data)} high-match jobs saved to '{csv_filename}'.")
        print(f"📄 Tailored Word documents are waiting for you in the 'Tailored_CVs' folder!")
        print(f"✉️  Tailored cover letters are waiting for you in the 'Cover_Letters' folder!")
        print(f"🌐 Digest report ready: {report_path} (open it in your browser)")
    else:
        print("\n⚠️ No jobs with a match score above 29 were found.")

if __name__ == "__main__":
    run_pipeline()
