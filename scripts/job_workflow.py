from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import shutil
import subprocess
import textwrap
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import quote_plus, urljoin

import httpx
import yaml
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright
from playwright.sync_api import Error as PlaywrightError
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
SETTINGS_FILE = ROOT / "config" / "settings.yaml"
PROFILE_FILE = ROOT / "config" / "candidate_profile.json"
JOBS_FILE = ROOT / "jobs" / "jobs.csv"
OUTPUT_APPLICATION_DIR = ROOT / "outputs" / "applications"
PROFILE_DIR = ROOT / "browser_profiles"
AUDIT_XLSX_FILE = ROOT / "application_audit.xlsx"


@dataclass
class JobRecord:
    job_id: str
    platform: str
    company: str
    title: str
    location: str
    job_url: str
    jd_text: str
    notes: str


def read_settings() -> dict:
    with SETTINGS_FILE.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_candidate_profile() -> dict:
    default_profile = {
        "name": "Douglas Mo",
        "email": "d157156@gmail.com",
        "phone": "0434 001 262",
        "linkedin": "https://www.linkedin.com/in/douglas-mo-67344531b/",
        "github": "https://github.com/DouglasMooooo",
        "portfolio": "https://douglasmo.vercel.app/",
        "gender": "Male",
        "ethnicity": "Chinese",
        "languages": ["English", "Mandarin", "Cantonese"],
        "address_line": "Algester Rd, Brisbane, QLD",
        "has_disability": "No",
        "is_lgbti": "No",
        "availability_note": "Available to start immediately for part-time or internship roles. Full-time from July 2026.",
    }
    if PROFILE_FILE.exists():
        try:
            with PROFILE_FILE.open("r", encoding="utf-8") as f:
                user_profile = json.load(f)
            if isinstance(user_profile, dict):
                default_profile.update(user_profile)
        except Exception:
            pass
    return default_profile


def get_verified_fact_snippets(profile: dict) -> List[str]:
    facts = []
    for k in ["name", "title", "location", "email", "phone", "work_authorization", "availability_note"]:
        v = str(profile.get(k, "")).strip()
        if v:
            facts.append(v)
    for item in profile.get("headline_achievements", []) or []:
        item = str(item).strip()
        if item:
            facts.append(item)
    return facts


def ensure_dirs() -> None:
    for path in [
        ROOT / "resumes",
        ROOT / "jobs",
        ROOT / "outputs",
        OUTPUT_APPLICATION_DIR,
        ROOT / "outputs" / "logs",
        PROFILE_DIR,
    ]:
        path.mkdir(parents=True, exist_ok=True)


def init_files() -> None:
    ensure_dirs()

    template_1 = ROOT / "resumes" / "resume_template_1.md"
    template_2 = ROOT / "resumes" / "resume_template_2.md"
    jobs_csv = JOBS_FILE

    if not template_1.exists():
        template_1.write_text(
            textwrap.dedent(
                """\
                # {name}
                Email: {email} | Phone: {phone} | LinkedIn: {linkedin}

                ## Professional Summary
                [Replace with your summary]

                ## Core Skills
                - [Skill A]
                - [Skill B]
                - [Skill C]

                ## Experience
                ### [Company] - [Role]
                - [Impact bullet with numbers]
                - [Impact bullet with numbers]

                ## Education
                - [Degree], [School]

                ## Certifications
                - [Certification]
                """
            ),
            encoding="utf-8",
        )

    if not template_2.exists():
        template_2.write_text(
            textwrap.dedent(
                """\
                # {name}
                {phone} | {email} | {city}

                ## Profile
                [Short profile for alternative role track]

                ## Technical Stack
                - [Tool/Language]
                - [Tool/Language]

                ## Selected Projects
                ### [Project Name]
                - [What was built]
                - [Business value]

                ## Work History
                ### [Company] - [Role]
                - [Result with KPI]

                ## Education
                - [Degree], [University]
                """
            ),
            encoding="utf-8",
        )

    if not jobs_csv.exists():
        jobs_csv.write_text(
            "job_id,platform,company,title,location,job_url,jd_text,notes\n",
            encoding="utf-8",
        )


def load_jobs() -> List[JobRecord]:
    if not JOBS_FILE.exists():
        raise FileNotFoundError(f"Missing jobs file: {JOBS_FILE}")

    rows: List[JobRecord] = []
    with JOBS_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(
                JobRecord(
                    job_id=(r.get("job_id") or "").strip(),
                    platform=(r.get("platform") or "").strip().lower(),
                    company=(r.get("company") or "").strip(),
                    title=(r.get("title") or "").strip(),
                    location=(r.get("location") or "").strip(),
                    job_url=(r.get("job_url") or "").strip(),
                    jd_text=(r.get("jd_text") or "").strip(),
                    notes=(r.get("notes") or "").strip(),
                )
            )
    return rows


def load_templates(settings: dict) -> List[Tuple[str, str, Path]]:
    templates: List[Tuple[str, str, Path]] = []
    for rel in settings.get("template_files", []):
        p = ROOT / rel
        if p.exists():
            text = read_template_text(p)
            if text.strip():
                templates.append((p.name, text, p))
    if not templates:
        raise FileNotFoundError("No templates found under resumes/")
    return templates


def read_template_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".md", ".txt"}:
        return path.read_text(encoding="utf-8")
    if suffix == ".docx":
        doc = Document(str(path))
        lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
        return "\n".join(lines)
    return ""


def extract_text_from_url(url: str, timeout_s: float = 20.0) -> str:
    if not url:
        return ""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"
        )
    }
    try:
        with httpx.Client(timeout=timeout_s, follow_redirects=True, headers=headers) as client:
            resp = client.get(url)
            resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "noscript"]):
            tag.extract()
        text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
        return text[:20000]
    except Exception:
        return ""


def login_platform(
    platform: str,
    headless: bool = False,
    wait_seconds: int = 120,
    channel: str = "msedge",
    use_system_edge: bool = False,
    cdp_port: int = 9222,
) -> None:
    platform = platform.lower()
    if platform not in {"seek", "linkedin"}:
        raise ValueError("platform must be seek or linkedin")

    profile_path = PROFILE_DIR / platform
    profile_path.mkdir(parents=True, exist_ok=True)
    login_url = "https://www.seek.com.au/" if platform == "seek" else "https://www.linkedin.com/"

    if use_system_edge:
        edge_exe = resolve_edge_executable()
        # Use user's installed Edge/profile via CDP, not Playwright isolated profile.
        subprocess.Popen(
            [
                edge_exe,
                f"--remote-debugging-port={cdp_port}",
                login_url,
            ],
            shell=False,
        )
        print(f"[{platform}] System Edge launched with CDP port {cdp_port}.")
        print(f"Please login in your normal Edge profile within {wait_seconds} seconds.")
        time.sleep(wait_seconds)
        print(f"[{platform}] System Edge login window completed.")
        return

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            headless=headless,
            viewport={"width": 1440, "height": 900},
            channel=channel,
            args=["--disable-blink-features=AutomationControlled"],
        )
        try:
            page = context.new_page()
            page.goto(login_url, wait_until="domcontentloaded")
            print(f"[{platform}] Browser opened. Please login manually.")
            print(f"Waiting {wait_seconds} seconds for manual login, then session will auto-save.")
            time.sleep(wait_seconds)
        except PlaywrightError:
            # User may close browser early; persistent profile may still contain login state.
            print(f"[{platform}] Browser closed before wait completed, attempting to save session.")
        try:
            context.storage_state(path=str(profile_path / "storage_state.json"))
        except PlaywrightError:
            pass
        try:
            context.close()
        except PlaywrightError:
            pass
        print(f"[{platform}] Session save attempt completed: {profile_path}")


def tokenize(text: str) -> List[str]:
    return re.findall(r"[A-Za-z][A-Za-z0-9\+\#\.\-]{1,}", text or "")


def keyword_stats(jd_text: str, resume_text: str, known_skills: List[str]) -> Dict[str, List[str]]:
    jd_tokens = {t.lower() for t in tokenize(jd_text)}
    resume_tokens = {t.lower() for t in tokenize(resume_text)}

    explicit = [k for k in known_skills if k.lower() in jd_tokens]
    covered = [k for k in explicit if k.lower() in resume_tokens]
    missing = [k for k in explicit if k.lower() not in resume_tokens]
    return {"covered": covered, "missing": missing}


def extract_top_jd_keywords(jd_text: str, limit: int = 24) -> List[str]:
    # Pull frequent longer tokens as ATS-oriented anchors.
    stop = {
        "with", "that", "this", "have", "will", "from", "your", "you", "our", "for",
        "and", "the", "are", "job", "role", "work", "team", "years", "year", "using",
        "experience", "skills", "ability", "required", "preferred", "about", "their",
        "into", "across", "through", "including", "within", "candidate",
    }
    tokens = [t.lower() for t in tokenize(jd_text) if len(t) >= 4]
    freq: Dict[str, int] = {}
    for t in tokens:
        if t in stop:
            continue
        freq[t] = freq.get(t, 0) + 1
    ranked = sorted(freq.items(), key=lambda x: (-x[1], x[0]))
    return [k for k, _ in ranked[:limit]]


def pick_template(templates: List[Tuple[str, str, Path]], jd_text: str) -> Tuple[str, str, Path]:
    jd_tokens = {t.lower() for t in tokenize(jd_text)}
    best_name = templates[0][0]
    best_text = templates[0][1]
    best_path = templates[0][2]
    best_score = -1
    for name, text, path in templates:
        tokens = {t.lower() for t in tokenize(text)}
        score = len(jd_tokens & tokens)
        if score > best_score:
            best_name, best_text, best_path, best_score = name, text, path, score
    return best_name, best_text, best_path


def build_resume_without_llm(
    template_text: str,
    job: JobRecord,
    covered: List[str],
    missing: List[str],
    profile: dict,
    ats_keywords: List[str],
) -> str:
    name = profile.get("name", "Douglas Mo")
    title = profile.get("title", "Business Analyst")
    contact = f"{profile.get('location','Brisbane, Australia')} | {profile.get('email','')} | {profile.get('phone','')}"
    top_kw = ", ".join(ats_keywords[:12]) if ats_keywords else ", ".join(covered[:10])
    achievement_lines = profile.get("headline_achievements", []) or []
    while len(achievement_lines) < 3:
        achievement_lines.append("Delivered measurable business outcomes through data-driven analysis and execution.")
    return textwrap.dedent(
        f"""\
        {name}
        {title}
        {contact}
        LinkedIn | GitHub | Portfolio

        SUMMARY
        Candidate aligned to {job.title} with strengths in {top_kw}. Combines customer-facing execution, business analysis, and commercial reporting to deliver measurable outcomes.
        {profile.get("availability_note", "")}

        SKILLS
        Business & Customer Operations
        - Customer service, customer experience, cross-functional collaboration, stakeholder communication, time management, problem solving
        Analytics & Reporting
        - KPI tracking, forecasting, variance analysis, performance reporting, data analysis, process improvement
        Tools
        - Microsoft Office, Excel, Power BI, Python, SQL, ERP/CRM exposure
        Role Keywords
        - {top_kw}

        EXPERIENCE
        Business / Operations Impact
        - Led analysis and reporting workflows that surfaced at-risk opportunities and improved decision quality across business teams.
        - {achievement_lines[0]}
        - {achievement_lines[1]}
        - {achievement_lines[2]}
        Retail Operations
        - Executed merchandising, stock replenishment, and transaction support in high-volume retail settings to improve customer outcomes and inventory reliability.
        AI & Workflow Improvement
        - Built AI-enabled workflow tools and automation that improved information retrieval speed and operational efficiency for internal users.

        EDUCATION
        Master of Business Analytics | Victoria University | 2024 - 2026
        Bachelor of Management (Financial Accounting) | Lingnan Normal University | 2018 - 2022
        """
    )


def build_cover_letter_without_llm(job: JobRecord, covered: List[str], missing: List[str], profile: dict) -> str:
    matched = ", ".join(covered[:8]) if covered else "relevant skills from the job description"
    gap_line = ", ".join(missing[:6]) if missing else "No obvious keyword gaps detected."
    verified_impact = (profile.get("headline_achievements") or [])[:2]
    impact_line = " ".join(verified_impact) if verified_impact else ""
    return textwrap.dedent(
        f"""\
        {datetime.now().strftime("%Y-%m-%d")}

        Hiring Manager
        {job.company}

        Re: {job.title}

        Dear Hiring Manager,

        I am writing to apply for the {job.title} opportunity at {job.company}. My background aligns strongly with the requirements in your job description, especially around {matched}.

        I focus on delivering measurable outcomes, clear communication, and reliable execution in cross-functional teams. I am confident I can add value quickly and contribute to your team in {job.location or "this role"}.

        Based on the role requirements, I have tailored my resume to emphasize the most relevant strengths while keeping claims accurate and evidence-based. {impact_line}

        {profile.get("availability_note", "")}

        Thank you for your time and consideration.

        Sincerely,
        Douglas Mo

        ---
        ATS alignment notes (internal):
        Missing keywords to consider adding in future updates: {gap_line}
        """
    )


def call_openai(prompt: str, model: str, api_key: str) -> str:
    # Uses Responses API directly to avoid SDK/version mismatch in local environments.
    url = "https://api.openai.com/v1/responses"
    payload = {
        "model": model,
        "input": prompt,
        "temperature": 0.2,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    with httpx.Client(timeout=120) as client:
        r = client.post(url, headers=headers, content=json.dumps(payload))
        r.raise_for_status()
        data = r.json()
    # Compatible fallback extraction:
    output_text = data.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()
    for item in data.get("output", []):
        for chunk in item.get("content", []):
            if chunk.get("type") == "output_text":
                txt = chunk.get("text", "")
                if txt.strip():
                    return txt.strip()
    return ""


def parse_resume_bundle(raw_text: str) -> dict:
    text = (raw_text or "").strip()
    if not text:
        return {}
    # Try direct JSON
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    # Try fenced JSON
    m = re.search(r"```json\s*(\{.*?\})\s*```", text, flags=re.DOTALL | re.IGNORECASE)
    if m:
        try:
            data = json.loads(m.group(1))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    # Try first object-like block
    m2 = re.search(r"(\{[\s\S]*\})", text)
    if m2:
        try:
            data = json.loads(m2.group(1))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}


def build_prompt(template_text: str, jd_text: str, job: JobRecord, ats_keywords: List[str]) -> str:
    return textwrap.dedent(
        f"""\
        You are a senior resume strategist and ATS optimization expert.
        Your goal is to REBUILD the resume to maximize interview probability.

        Follow this strict system:
        STEP 1 ROLE LOCK:
        - Identify ONE core role and align the entire resume to only that role.
        - Remove irrelevant content.

        STEP 2 KEYWORD DOMINATION:
        - Extract 20-30 keywords from JD.
        - Force-inject keywords into Summary (3-5), Skills (>=12), Experience (>=10).

        STEP 3 IMPACT REWRITE:
        - Rewrite bullets using Action + Business Impact + Metric.
        - Show value creation, problem solving, measurable outcomes.

        STEP 4 SIMPLIFICATION:
        - Remove redundant wording and unnecessary jargon.
        - Make it readable in under 10 seconds.

        STEP 5 DUAL OUTPUT + VARIATIONS:
        - Produce three resume variations:
          1) safe_version: high ATS pass rate
          2) aggressive_version: maximum keyword + impact
          3) human_version: best readability and persuasion

        STEP 6 SCORING:
        - ATS match score (0-100)
        - Missing keywords
        - Risk flags

        STEP 7 IMPROVEMENT LOOP:
        - 3 improvements to further increase match score
        - what to remove for clarity

        HARD RULES:
        - No fabricated experience, no fake companies, dates, or metrics.
        - Standard headings: SUMMARY, SKILLS, EXPERIENCE, EDUCATION.
        - No tables/graphics/icons.
        - Copy-ready output only.

        Return STRICT JSON only with this schema:
        {{
          "role_lock": "string",
          "keywords_injected": ["..."],
          "ats_match_score": 0,
          "missing_keywords": ["..."],
          "risk_flags": ["..."],
          "improvements": ["..."],
          "remove_for_clarity": ["..."],
          "safe_version": "full resume text",
          "aggressive_version": "full resume text",
          "human_version": "full resume text"
        }}

        Job meta:
        - Title: {job.title}
        - Company: {job.company}
        - Location: {job.location}
        - ATS target keywords: {", ".join(ats_keywords[:30])}

        JD:
        {jd_text[:12000]}

        Base resume:
        {template_text}
        """
    )


def build_cover_letter_prompt(job: JobRecord, jd_text: str, resume_text: str, ats_keywords: List[str]) -> str:
    return textwrap.dedent(
        f"""\
        You are a professional career writer.
        Write a tailored cover letter in markdown for this exact job.

        Constraints:
        - Keep it concise (220-320 words).
        - Tone: confident, specific, professional.
        - Use ATS-relevant wording naturally from JD.
        - No invented facts, achievements, or technologies.
        - Use this structure:
          1) opening with role + motivation
          2) 1-2 paragraphs on fit and impact
          3) closing + call to action
        - Do not include placeholders.

        Job meta:
        - Title: {job.title}
        - Company: {job.company}
        - Location: {job.location}
        - ATS target keywords: {", ".join(ats_keywords[:20])}

        JD:
        {jd_text[:12000]}

        Resume:
        {resume_text[:12000]}
        """
    )


def evaluate_resume(jd_text: str, resume_text: str, known_skills: List[str]) -> Dict[str, str]:
    lower_resume = resume_text.lower()
    ats_keywords = extract_top_jd_keywords(jd_text, limit=24)
    resume_tokens = {t.lower() for t in tokenize(resume_text)}
    ats_matched = [k for k in ats_keywords if k.lower() in resume_tokens]
    missing = [k for k in ats_keywords if k.lower() not in resume_tokens]

    # 1) ATS keyword coverage (55%)
    kw_score = round(100 * len(ats_matched) / max(len(ats_keywords), 1), 1)

    # 2) Standard sections (15%)
    section_targets = ["summary", "skills", "experience", "education"]
    present = sum(1 for s in section_targets if s in lower_resume)
    section_score = round(100 * present / len(section_targets), 1)

    # 3) Impact density: bullets with measurable outcomes (20%)
    bullets = [ln.strip() for ln in resume_text.splitlines() if ln.strip().startswith("-")]
    measurable = 0
    for b in bullets:
        if re.search(r"(\d+|%|\$|million|kpi|roi|latency|accuracy|saved|reduced|improved)", b.lower()):
            measurable += 1
    impact_score = round(100 * measurable / max(len(bullets), 1), 1)

    # 4) Role-title alignment (10%)
    jd_title_tokens = [t.lower() for t in tokenize(jd_text[:220]) if len(t) >= 4][:10]
    role_hits = sum(1 for t in jd_title_tokens if t in lower_resume)
    role_score = round(100 * role_hits / max(len(jd_title_tokens), 1), 1)

    ats_score = round(
        kw_score * 0.55 + section_score * 0.15 + impact_score * 0.20 + role_score * 0.10,
        1,
    )
    return {
        "match_score": f"{ats_score}",
        "ats_score": f"{ats_score}",
        "section_score": f"{section_score}",
        "ats_keywords_used": ", ".join(ats_matched[:20]) if ats_matched else "",
        "covered_keywords": ", ".join(ats_matched[:20]) if ats_matched else "",
        "missing_keywords": ", ".join(missing[:15]) if missing else "None",
    }


def run_pipeline(use_openai: bool) -> None:
    ensure_dirs()
    load_dotenv(ROOT / ".env")
    settings = read_settings()
    templates = load_templates(settings)
    jobs = load_jobs()
    known_skills = settings.get("known_skills", [])
    candidate_profile = load_candidate_profile()

    model = os.getenv("OPENAI_MODEL", settings.get("default_openai_model", "gpt-5-mini"))
    api_key = os.getenv("OPENAI_API_KEY", "")
    can_use_openai = use_openai and bool(api_key)

    results = []
    for i, job in enumerate(jobs, start=1):
        job_id = job.job_id or f"job_{i:03d}"
        jd_text = job.jd_text.strip()
        if not jd_text and job.job_url:
            jd_text = extract_text_from_url(job.job_url)

        if not jd_text:
            results.append(
                {
                    "job_id": job_id,
                    "platform": job.platform,
                    "company": job.company,
                    "title": job.title,
                    "status": "skipped_no_jd",
                    "resume_file": "",
                    "resume_safe_file": "",
                    "resume_aggressive_file": "",
                    "resume_human_file": "",
                    "cover_letter_file": "",
                    "role_lock": "",
                    "keywords_injected": "",
                    "match_score": "",
                    "ats_match_score_llm": "",
                    "ats_score": "",
                    "section_score": "",
                    "ats_keywords_used": "",
                    "covered_keywords": "",
                    "missing_keywords": "",
                    "missing_keywords_llm": "",
                    "risk_flags": "",
                    "improvements": "",
                    "remove_for_clarity": "",
                    "job_url": job.job_url,
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                }
            )
            continue

        template_name, template_text, template_path = pick_template(templates, jd_text)
        stats = keyword_stats(jd_text, template_text, known_skills)
        ats_keywords = extract_top_jd_keywords(jd_text)
        template_links = extract_docx_field_links(template_path) if template_path.suffix.lower() == ".docx" else {}

        safe_resume_text = ""
        aggressive_resume_text = ""
        human_resume_text = ""
        role_lock = ""
        injected_keywords = []
        llm_ats_match_score = ""
        llm_missing_keywords = ""
        llm_risk_flags = ""
        llm_improvements = ""
        llm_remove_for_clarity = ""

        if can_use_openai:
            prompt = build_prompt(template_text, jd_text, job, ats_keywords)
            try:
                llm_raw = call_openai(prompt, model=model, api_key=api_key)
                bundle = parse_resume_bundle(llm_raw)
                safe_resume_text = str(bundle.get("safe_version", "")).strip()
                aggressive_resume_text = str(bundle.get("aggressive_version", "")).strip()
                human_resume_text = str(bundle.get("human_version", "")).strip()
                role_lock = str(bundle.get("role_lock", "")).strip()
                injected_keywords = bundle.get("keywords_injected", []) or []
                llm_ats_match_score = str(bundle.get("ats_match_score", "")).strip()
                llm_missing_keywords = ", ".join(bundle.get("missing_keywords", []) or [])
                llm_risk_flags = ", ".join(bundle.get("risk_flags", []) or [])
                llm_improvements = " | ".join(bundle.get("improvements", []) or [])
                llm_remove_for_clarity = ", ".join(bundle.get("remove_for_clarity", []) or [])
                if not human_resume_text:
                    raise ValueError("Missing human_version in LLM JSON output")
            except Exception:
                human_resume_text = build_resume_without_llm(
                    template_text, job, stats["covered"], stats["missing"], candidate_profile, ats_keywords
                )
                safe_resume_text = human_resume_text
                aggressive_resume_text = human_resume_text
        else:
            human_resume_text = build_resume_without_llm(
                template_text, job, stats["covered"], stats["missing"], candidate_profile, ats_keywords
            )
            safe_resume_text = human_resume_text
            aggressive_resume_text = human_resume_text

        if can_use_openai:
            cl_prompt = build_cover_letter_prompt(job, jd_text, human_resume_text, ats_keywords)
            try:
                cover_letter_text = call_openai(cl_prompt, model=model, api_key=api_key)
                if not cover_letter_text.strip():
                    raise ValueError("Empty OpenAI output")
            except Exception:
                cover_letter_text = build_cover_letter_without_llm(
                    job, stats["covered"], stats["missing"], candidate_profile
                )
        else:
            cover_letter_text = build_cover_letter_without_llm(
                job, stats["covered"], stats["missing"], candidate_profile
            )

        evaluation = evaluate_resume(jd_text, human_resume_text, known_skills)
        out_file, cl_file = build_output_paths(job.company, job_id)
        out_file = write_resume_docx(human_resume_text, out_file, template_links, candidate_profile)
        cl_file = write_cover_letter_docx(cover_letter_text, cl_file)

        results.append(
            {
                "job_id": job_id,
                "platform": job.platform,
                "company": job.company,
                "title": job.title,
                "status": "ready_to_apply",
                "resume_file": str(out_file),
                "resume_safe_file": "",
                "resume_aggressive_file": "",
                "resume_human_file": str(out_file),
                "cover_letter_file": str(cl_file),
                "template_used": template_name,
                "role_lock": role_lock,
                "keywords_injected": ", ".join(injected_keywords) if isinstance(injected_keywords, list) else str(injected_keywords),
                "match_score": evaluation["match_score"],
                "ats_match_score_llm": llm_ats_match_score,
                "ats_score": evaluation["ats_score"],
                "section_score": evaluation["section_score"],
                "ats_keywords_used": evaluation["ats_keywords_used"],
                "covered_keywords": evaluation["covered_keywords"],
                "missing_keywords": evaluation["missing_keywords"],
                "missing_keywords_llm": llm_missing_keywords,
                "risk_flags": llm_risk_flags,
                "improvements": llm_improvements,
                "remove_for_clarity": llm_remove_for_clarity,
                "strict_profile_facts_used": " | ".join(get_verified_fact_snippets(candidate_profile)[:8]),
                "job_url": job.job_url,
                "timestamp": datetime.now().isoformat(timespec="seconds"),
            }
        )

    write_xlsx(results, AUDIT_XLSX_FILE)
    print(f"Updated audit file: {AUDIT_XLSX_FILE}")


def safe_name(s: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9\-_]+", "_", (s or "role")).strip("_")
    return cleaned[:60] if cleaned else "role"


def build_output_paths(company: str, job_id: str) -> Tuple[Path, Path]:
    base = safe_name(company) or "company"
    resume = OUTPUT_APPLICATION_DIR / f"{base}_resume.docx"
    cover = OUTPUT_APPLICATION_DIR / f"{base}_coverletter.docx"
    if resume.exists() or cover.exists():
        resume = OUTPUT_APPLICATION_DIR / f"{base}_{safe_name(job_id)}_resume.docx"
        cover = OUTPUT_APPLICATION_DIR / f"{base}_{safe_name(job_id)}_coverletter.docx"
    return resume, cover


def add_hyperlink(paragraph, text: str, url: str) -> None:
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1155CC")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(color)
    r_pr.append(underline)
    run.append(r_pr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def extract_docx_field_links(path: Path) -> Dict[str, str]:
    links: Dict[str, str] = {}
    try:
        with zipfile.ZipFile(path, "r") as zf:
            xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    except Exception:
        return links

    field_pattern = re.compile(
        r'HYPERLINK\s+&quot;([^&]+)&quot;(.+?)w:fldChar w:fldCharType="end"',
        flags=re.DOTALL,
    )
    for m in field_pattern.finditer(xml):
        url = html.unescape(m.group(1)).strip()
        seg = m.group(2)
        texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", seg, flags=re.DOTALL)
        label = html.unescape("".join(texts)).strip()
        if not label or not url:
            continue
        links[label.lower()] = url
    return links


def style_doc_defaults(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)


def write_resume_docx(text: str, path: Path, template_links: Dict[str, str], profile: Optional[dict] = None) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    style_doc_defaults(doc)

    section_headers = {
        "SUMMARY",
        "SKILLS",
        "EXPERIENCE",
        "EDUCATION",
        "PROJECTS",
        "CERTIFICATIONS",
        "AVAILABILITY",
    }

    raw_lines = text.splitlines()
    lines: List[str] = []
    for ln in raw_lines:
        t = ln.strip()
        if t.startswith("<!--") and t.endswith("-->"):
            continue
        lines.append(ln)

    i = 0
    if lines:
        first = doc.add_paragraph(lines[0].strip())
        if first.runs:
            first.runs[0].bold = True
            first.runs[0].font.size = Pt(16)
        i = 1
    if i < len(lines):
        second = lines[i].strip()
        if second:
            p2 = doc.add_paragraph(second)
            if p2.runs:
                p2.runs[0].bold = True
                p2.runs[0].font.size = Pt(11)
        i += 1
    if i < len(lines):
        third = lines[i].strip()
        if third:
            doc.add_paragraph(third)
        i += 1

    if template_links:
        p = doc.add_paragraph("🔗 ")
        add_hyperlink(p, "LinkedIn", template_links.get("linkedin", "https://www.linkedin.com/"))
        p.add_run(" | 💻 ")
        add_hyperlink(p, "GitHub", template_links.get("github", "https://github.com/"))
        p.add_run(" | 🌐 ")
        add_hyperlink(p, "Portfolio", template_links.get("portfolio", "https://"))
    elif profile:
        p = doc.add_paragraph("🔗 ")
        add_hyperlink(p, "LinkedIn", profile.get("linkedin", "https://www.linkedin.com/"))
        p.add_run(" | 💻 ")
        add_hyperlink(p, "GitHub", profile.get("github", "https://github.com/"))
        p.add_run(" | 🌐 ")
        add_hyperlink(p, "Portfolio", profile.get("portfolio", "https://"))

    for raw in lines[i:]:
        line = raw.rstrip()
        if not line.strip():
            doc.add_paragraph("")
            continue

        clean = line.strip()
        low = clean.lower()
        if "linkedin" in low and "github" in low and "portfolio" in low:
            # We already render a clickable hyperlink row above.
            continue
        upper = clean.upper()
        if upper in section_headers:
            p = doc.add_paragraph(clean)
            if p.runs:
                p.runs[0].bold = True
                p.runs[0].font.size = Pt(11.5)
            continue

        if clean.startswith("### ") or clean.startswith("## ") or clean.startswith("# "):
            title = clean.lstrip("# ").strip()
            p = doc.add_paragraph(title)
            if p.runs:
                p.runs[0].bold = True
            continue

        if clean.startswith("- "):
            p = doc.add_paragraph(clean[2:].strip(), style="List Bullet")
            continue

        # Emphasize likely role/company lines for readability.
        if "|" in clean and ("–" in clean or "-" in clean):
            p = doc.add_paragraph(clean)
            if p.runs:
                p.runs[0].bold = True
            continue

        doc.add_paragraph(clean)

    return save_doc_with_fallback(doc, path)


def write_cover_letter_docx(text: str, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    style_doc_defaults(doc)
    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            doc.add_paragraph("")
            continue
        if line.startswith("### "):
            p = doc.add_paragraph(line[4:].strip())
            p.style = "Heading 3"
            continue
        if line.startswith("## "):
            p = doc.add_paragraph(line[3:].strip())
            p.style = "Heading 2"
            continue
        if line.startswith("# "):
            p = doc.add_paragraph(line[2:].strip())
            p.style = "Heading 1"
            continue
        if line.startswith("- "):
            p = doc.add_paragraph(line[2:].strip())
            try:
                p.style = "List Bullet"
            except Exception:
                pass
            continue
        doc.add_paragraph(line)
    return save_doc_with_fallback(doc, path)


def save_doc_with_fallback(doc: Document, path: Path) -> Path:
    try:
        doc.save(path)
        return path
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = path.with_name(f"{path.stem}_{ts}{path.suffix}")
        doc.save(alt)
        return alt


def resolve_edge_executable() -> str:
    found = shutil.which("msedge")
    if found:
        return found
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    raise FileNotFoundError("Microsoft Edge executable not found.")


def write_csv(rows: List[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        headers = [
            "job_id",
            "platform",
            "company",
            "title",
            "status",
            "resume_file",
            "resume_safe_file",
            "resume_aggressive_file",
            "resume_human_file",
            "cover_letter_file",
            "template_used",
            "role_lock",
            "keywords_injected",
            "match_score",
            "ats_match_score_llm",
            "ats_score",
            "section_score",
            "ats_keywords_used",
            "covered_keywords",
            "missing_keywords",
            "missing_keywords_llm",
            "risk_flags",
            "improvements",
            "remove_for_clarity",
            "strict_profile_facts_used",
            "job_url",
            "timestamp",
            "applied_at",
            "auto_step",
        ]
    else:
        headers = list(rows[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def write_xlsx(rows: List[dict], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "results"
    if not rows:
        ws.append(
            [
                "job_id",
                "platform",
                "company",
                "title",
                "status",
                "resume_file",
                "resume_safe_file",
                "resume_aggressive_file",
                "resume_human_file",
                "cover_letter_file",
                "template_used",
                "role_lock",
                "keywords_injected",
                "match_score",
                "ats_match_score_llm",
                "ats_score",
                "section_score",
                "ats_keywords_used",
                "covered_keywords",
                "missing_keywords",
                "missing_keywords_llm",
                "risk_flags",
                "improvements",
                "remove_for_clarity",
                "strict_profile_facts_used",
                "job_url",
                "timestamp",
                "applied_at",
                "auto_step",
            ]
        )
    else:
        headers = list(rows[0].keys())
        ws.append(headers)
        for row in rows:
            ws.append([row.get(h, "") for h in headers])
    wb.save(path)


def read_xlsx_rows(path: Path) -> List[dict]:
    if not path.exists():
        return []
    wb = load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h) if h is not None else "" for h in rows[0]]
    data: List[dict] = []
    for r in rows[1:]:
        row_dict = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            v = r[i] if i < len(r) else ""
            row_dict[h] = "" if v is None else str(v)
        if row_dict:
            data.append(row_dict)
    return data


def assist_apply(platform: str, limit: int = 20, headless: bool = False) -> None:
    if not AUDIT_XLSX_FILE.exists():
        raise FileNotFoundError(f"Audit file not found: {AUDIT_XLSX_FILE}. Run pipeline first.")

    rows = read_xlsx_rows(AUDIT_XLSX_FILE)
    pending_pairs: List[Tuple[int, dict]] = []
    for idx, row in enumerate(rows):
        if row.get("platform", "").lower() == platform.lower() and row.get("status", "") == "ready_to_apply":
            pending_pairs.append((idx, row))
    pending_pairs = pending_pairs[:limit]

    if not pending_pairs:
        print(f"No pending rows for platform={platform}")
        return

    profile_path = PROFILE_DIR / platform.lower()
    if not profile_path.exists():
        raise FileNotFoundError(
            f"Missing browser profile for {platform}. Run auth first: "
            f"py scripts/job_workflow.py auth --platform {platform}"
        )

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            headless=headless,
            viewport={"width": 1440, "height": 900},
        )
        page = context.new_page()
        for idx, row in pending_pairs:
            url = str(row.get("job_url", "")).strip()
            if not url:
                rows[idx]["status"] = "skipped_no_url"
                continue

            print("=" * 60)
            print(f"Opening: {row.get('title', '')} @ {row.get('company', '')}")
            print(f"Resume: {row.get('resume_file', '')}")
            try:
                page.goto(url, wait_until="domcontentloaded")
            except PlaywrightTimeoutError:
                pass
            print("Please complete application manually in browser.")
            print("Type one of: submitted / skipped / failed")
            decision = input("> ").strip().lower() or "skipped"
            if decision not in {"submitted", "skipped", "failed"}:
                decision = "skipped"
            rows[idx]["status"] = decision
            rows[idx]["applied_at"] = datetime.now().isoformat(timespec="seconds")
        context.close()

    write_xlsx(rows, AUDIT_XLSX_FILE)
    print(f"Audit updated: {AUDIT_XLSX_FILE}")


def click_first(page, selectors: List[str]) -> bool:
    scopes = [page] + list(page.frames)
    for scope in scopes:
        for sel in selectors:
            try:
                loc = scope.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=1200):
                    loc.click(timeout=2200)
                    return True
            except Exception:
                continue
    return False


def try_auto_click_apply(page, platform: str) -> str:
    platform = platform.lower()
    if platform == "seek":
        selectors = [
            "button:has-text('Apply')",
            "button:has-text('Quick apply')",
            "a:has-text('Apply')",
            "[data-automation*='apply']",
        ]
    else:
        selectors = [
            "button:has-text('Easy Apply')",
            "button:has-text('Apply')",
            "a:has-text('Easy Apply')",
            "[aria-label*='Easy Apply']",
        ]
    ok = click_first(page, selectors)
    return "auto_clicked_apply" if ok else "auto_click_not_found"


def click_apply_chain(page) -> List[str]:
    actions: List[str] = []
    selectors = [
        "button:has-text('Apply')",
        "button:has-text('Apply now')",
        "button:has-text('Quick apply')",
        "button:has-text('Easy Apply')",
        "a:has-text('Apply')",
        "a:has-text('Apply now')",
        "a:has-text('Apply on company site')",
        "a:has-text('Apply on employer site')",
        "a:has-text('Continue')",
        "button:has-text('Continue')",
    ]
    for _ in range(4):
        clicked = click_first(page, selectors)
        if not clicked:
            break
        actions.append("clicked_apply_step")
        page.wait_for_timeout(1200)
    return actions


def fill_first(page, selectors: List[str], value: str) -> bool:
    if not value:
        return False
    scopes = [page] + list(page.frames)
    for scope in scopes:
        for sel in selectors:
            try:
                loc = scope.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=900):
                    existing = ""
                    try:
                        existing = (loc.input_value(timeout=900) or "").strip()
                    except Exception:
                        existing = ""
                    # Strict safety: do not overwrite existing user/provided content.
                    if existing:
                        continue
                    loc.fill(value, timeout=2200)
                    return True
            except Exception:
                continue
    return False


def autofill_form(page, profile: dict, resume_file: str, cover_letter_text: str) -> Dict[str, int]:
    # Strict whitelist-only autofill to avoid hallucinated form mappings.
    stats = {"fields": 0, "uploads": 0, "coverletter": 0}
    mappings = {
        profile.get("name", "").split(" ")[0]: [
            "input[name*='first']",
            "input[id*='first']",
            "input[placeholder*='First']",
        ],
        " ".join(profile.get("name", "").split(" ")[1:]): [
            "input[name*='last']",
            "input[id*='last']",
            "input[placeholder*='Last']",
        ],
    }
    if fill_first(page, mappings.get(profile.get("name", "").split(" ")[0], []), profile.get("name", "").split(" ")[0]):
        stats["fields"] += 1
    last_name = " ".join(profile.get("name", "").split(" ")[1:]).strip()
    if last_name and fill_first(page, mappings.get(last_name, []), last_name):
        stats["fields"] += 1

    field_map = [
        (profile.get("email", ""), ["input[type='email']", "input[name*='email']", "input[id*='email']"]),
        (profile.get("phone", ""), ["input[type='tel']", "input[name*='phone']", "input[id*='phone']"]),
        (profile.get("address_line", ""), ["input[name*='address']", "input[id*='address']", "input[placeholder*='Address']"]),
        (profile.get("location", ""), ["input[name*='city']", "input[id*='city']", "input[placeholder*='City']"]),
    ]
    for value, selectors in field_map:
        if fill_first(page, selectors, value):
            stats["fields"] += 1

    # Select-based demographic fields where applicable.
    for value, selectors in [
        (profile.get("gender", ""), ["select[name*='gender']", "select[id*='gender']"]),
        (profile.get("has_disability", ""), ["select[name*='disab']", "select[id*='disab']"]),
        (profile.get("is_lgbti", ""), ["select[name*='lgbt']", "select[id*='lgbt']"]),
    ]:
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0:
                    ok = False
                    try:
                        loc.select_option(label=value, timeout=1500)
                        ok = True
                    except Exception:
                        try:
                            loc.select_option(value=value.lower(), timeout=1500)
                            ok = True
                        except Exception:
                            ok = False
                    if ok:
                        stats["fields"] += 1
                        break
            except Exception:
                continue

    # Upload CV
    try:
        file_input = page.locator("input[type='file']").first
        if file_input.count() > 0:
            file_input.set_input_files(resume_file, timeout=3000)
            stats["uploads"] += 1
    except Exception:
        pass

    # Cover letter textarea
    if cover_letter_text:
        if fill_first(
            page,
            ["textarea[name*='cover']", "textarea[id*='cover']", "textarea[placeholder*='cover']", "textarea[name*='letter']", "textarea[id*='letter']"],
            cover_letter_text[:3500],
        ):
            stats["coverletter"] += 1

    return stats


def read_cover_letter_text(path: str) -> str:
    p = Path(path)
    if not p.exists():
        return ""
    try:
        doc = Document(str(p))
        return "\n".join([x.text for x in doc.paragraphs if x.text.strip()])
    except Exception:
        return ""


def scrape_seek_jobs(
    query: str,
    location: str,
    limit: int = 20,
    channel: str = "msedge",
    headless: bool = False,
    use_system_edge: bool = False,
    cdp_url: str = "http://127.0.0.1:9222",
) -> List[dict]:
    def to_seek_slug(value: str) -> str:
        clean = re.sub(r"[^A-Za-z0-9 ]+", " ", (value or "").strip())
        clean = re.sub(r"\s+", "-", clean).strip("-")
        return clean or "jobs"

    def norm_txt(value: str) -> str:
        return re.sub(r"\s+", " ", (value or "")).strip()

    def extract_jobposting_from_ldjson(raw: str) -> Optional[dict]:
        try:
            payload = json.loads(raw)
        except Exception:
            return None
        stack = [payload]
        while stack:
            item = stack.pop()
            if isinstance(item, list):
                stack.extend(item)
                continue
            if not isinstance(item, dict):
                continue
            t = str(item.get("@type", "")).lower()
            if "jobposting" in t:
                return item
            for k in ("@graph", "graph", "itemListElement"):
                v = item.get(k)
                if isinstance(v, (list, dict)):
                    stack.append(v)
        return None

    def parse_seek_detail(page2, fallback_title: str, fallback_loc: str) -> dict:
        body = ""
        try:
            body = page2.inner_text("body")
        except Exception:
            body = ""
        title = ""
        company = ""
        loc2 = ""
        jd_text = body or ""
        try:
            title = norm_txt(page2.locator("h1").first.inner_text(timeout=1000))
        except Exception:
            title = ""
        try:
            company = norm_txt(
                page2.locator("a[data-automation='advertiser-name'], [data-automation='advertiser-name']").first.inner_text(timeout=1000)
            )
        except Exception:
            company = ""
        try:
            loc2 = norm_txt(page2.locator("[data-automation='job-detail-location']").first.inner_text(timeout=1000))
        except Exception:
            loc2 = ""
        try:
            jd = page2.locator("[data-automation='jobAdDetails']").first.inner_text(timeout=1200)
            if jd and len(jd.strip()) > 100:
                jd_text = jd
        except Exception:
            pass

        ld_scripts = []
        try:
            ld_scripts = page2.locator("script[type='application/ld+json']").all_text_contents()
        except Exception:
            ld_scripts = []
        for raw in ld_scripts:
            jobposting = extract_jobposting_from_ldjson(raw)
            if not jobposting:
                continue
            title = title or norm_txt(str(jobposting.get("title", "")))
            org = jobposting.get("hiringOrganization") or {}
            if isinstance(org, dict):
                company = company or norm_txt(str(org.get("name", "")))
            desc = jobposting.get("description")
            if isinstance(desc, str) and len(desc.strip()) > 50:
                desc_clean = norm_txt(re.sub(r"<[^>]+>", " ", html.unescape(desc)))
                if len(desc_clean) > 100:
                    jd_text = desc_clean
            job_loc = jobposting.get("jobLocation")
            if isinstance(job_loc, dict):
                job_loc = [job_loc]
            if isinstance(job_loc, list) and job_loc:
                addr = (job_loc[0] or {}).get("address", {})
                if isinstance(addr, dict):
                    loc_bits = [
                        str(addr.get("addressLocality", "")).strip(),
                        str(addr.get("addressRegion", "")).strip(),
                    ]
                    loc_joined = " ".join([x for x in loc_bits if x]).strip()
                    if loc_joined:
                        loc2 = loc2 or loc_joined
            break

        if not loc2:
            mloc = re.search(r"([A-Za-z ]+,?\s+QLD|Brisbane QLD)", body or "", flags=re.I)
            if mloc:
                loc2 = norm_txt(mloc.group(1))
        quick = ("quick apply" in (body or "").lower()) or ("easy apply" in (body or "").lower())
        return {
            "title": title or norm_txt(fallback_title) or "Unknown title",
            "company": company or "Unknown company",
            "location": loc2 or fallback_loc,
            "jd_text": (jd_text or "")[:12000],
            "quick_apply": quick,
            "notes": "quick_apply_detected" if quick else "standard_apply",
        }

    q = (query or "").strip() or "business graduate"
    loc = (location or "").strip() or "Brisbane QLD"
    url = f"https://www.seek.com.au/{to_seek_slug(q)}-jobs/in-{to_seek_slug(loc)}"
    jobs: List[dict] = []
    with sync_playwright() as p:
        browser = None
        context = None
        created_context = False
        if use_system_edge:
            browser = p.chromium.connect_over_cdp(cdp_url)
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            created_context = len(browser.contexts) == 0
        else:
            browser = p.chromium.launch(channel=channel, headless=headless)
            context = browser.new_context()
        page = context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3500)
        page.mouse.wheel(0, 2800)
        page.wait_for_timeout(1800)
        cards = page.locator("a[href*='/job/']").all()[: max(limit * 8, 80)]
        seen = set()
        for a in cards:
            try:
                href = a.get_attribute("href") or ""
                title = norm_txt(a.inner_text() or "")
            except Exception:
                continue
            if not href or "/job/" not in href:
                continue
            if href.startswith("/"):
                href = "https://www.seek.com.au" + href
            href = href.split("?")[0]
            if href in seen:
                continue
            seen.add(href)
            # Navigate briefly to collect more details.
            try:
                page2 = context.new_page()
                page2.goto(href, wait_until="domcontentloaded", timeout=45000)
                page2.wait_for_timeout(1400)
                parsed = parse_seek_detail(page2, fallback_title=title, fallback_loc=loc)
                jobs.append(
                    {
                        "job_url": href,
                        "title": parsed["title"],
                        "company": norm_txt(parsed["company"]),
                        "location": norm_txt(parsed["location"]),
                        "jd_text": parsed["jd_text"],
                        "notes": parsed["notes"],
                        "quick_apply": bool(parsed["quick_apply"]),
                    }
                )
                page2.close()
            except Exception:
                continue
            if len(jobs) >= limit:
                break
        if not use_system_edge or created_context:
            context.close()
        if not use_system_edge and browser is not None:
            browser.close()
    # Sort quick apply first for easier automation testing.
    jobs.sort(key=lambda x: (not x.get("quick_apply", False), x.get("title", "")))
    return jobs[:limit]


def scrape_linkedin_jobs(
    query: str,
    location: str,
    limit: int = 20,
    channel: str = "msedge",
    headless: bool = False,
    use_system_edge: bool = False,
    cdp_url: str = "http://127.0.0.1:9222",
) -> List[dict]:
    def norm_txt(value: str) -> str:
        return re.sub(r"\s+", " ", (value or "")).strip()

    def dedupe_repeated_phrase(value: str) -> str:
        txt = norm_txt(value)
        if not txt:
            return txt
        parts = txt.split(" ")
        n = len(parts)
        if n >= 4 and n % 2 == 0:
            mid = n // 2
            if " ".join(parts[:mid]).lower() == " ".join(parts[mid:]).lower():
                return " ".join(parts[:mid]).strip()
        return txt

    def extract_jobposting_from_ldjson(raw: str) -> Optional[dict]:
        try:
            payload = json.loads(raw)
        except Exception:
            return None
        stack = [payload]
        while stack:
            item = stack.pop()
            if isinstance(item, list):
                stack.extend(item)
                continue
            if not isinstance(item, dict):
                continue
            t = str(item.get("@type", "")).lower()
            if "jobposting" in t:
                return item
            for k in ("@graph", "graph", "itemListElement"):
                v = item.get(k)
                if isinstance(v, (list, dict)):
                    stack.append(v)
        return None

    def parse_linkedin_detail(page2, fallback_title: str, fallback_company: str, fallback_loc: str) -> dict:
        title = ""
        company = ""
        loc2 = ""
        jd_text = ""
        body = ""
        try:
            body = page2.inner_text("body")
        except Exception:
            body = ""
        for sel in ["h1.top-card-layout__title", "h1.t-24", "h1"]:
            try:
                title = norm_txt(page2.locator(sel).first.inner_text(timeout=900))
                if title:
                    break
            except Exception:
                continue
        for sel in [
            "a.topcard__org-name-link",
            ".topcard__flavor-row .topcard__flavor",
            "div.top-card-layout__card a",
            "span.jobs-unified-top-card__company-name",
        ]:
            try:
                company = norm_txt(page2.locator(sel).first.inner_text(timeout=900))
                if company:
                    break
            except Exception:
                continue
        # Fallback: parse from document title like "Role title | Company | LinkedIn".
        if not company:
            try:
                page_title = norm_txt(page2.title())
                chunks = [x.strip() for x in page_title.split("|") if x.strip()]
                if len(chunks) >= 2 and "linkedin" in chunks[-1].lower():
                    company = chunks[-2]
            except Exception:
                pass
        for sel in [
            "span.topcard__flavor--bullet",
            ".topcard__flavor.topcard__flavor--bullet",
            ".jobs-unified-top-card__bullet",
            "span.jobs-unified-top-card__workplace-type",
        ]:
            try:
                loc2 = norm_txt(page2.locator(sel).first.inner_text(timeout=900))
                if loc2:
                    break
            except Exception:
                continue
        for sel in [
            "div.show-more-less-html__markup",
            "div.jobs-description__content",
            "article.jobs-description",
            "section.description",
        ]:
            try:
                jd_text = norm_txt(page2.locator(sel).first.inner_text(timeout=1000))
                if jd_text:
                    break
            except Exception:
                continue

        ld_scripts = []
        try:
            ld_scripts = page2.locator("script[type='application/ld+json']").all_text_contents()
        except Exception:
            ld_scripts = []
        for raw in ld_scripts:
            jp = extract_jobposting_from_ldjson(raw)
            if not jp:
                continue
            title = title or norm_txt(str(jp.get("title", "")))
            org = jp.get("hiringOrganization") or {}
            if isinstance(org, dict):
                company = company or norm_txt(str(org.get("name", "")))
            desc = jp.get("description")
            if isinstance(desc, str) and len(desc.strip()) > 40:
                desc_clean = norm_txt(re.sub(r"<[^>]+>", " ", html.unescape(desc)))
                if len(desc_clean) > 100:
                    jd_text = jd_text or desc_clean
            break

        title = dedupe_repeated_phrase(title)
        if title:
            title = re.sub(r"\s+with verification$", "", title, flags=re.I).strip()
        easy = "easy apply" in (body or "").lower()
        return {
            "title": title or fallback_title or "LinkedIn role",
            "company": company or fallback_company or "Unknown company",
            "location": loc2 or fallback_loc,
            "jd_text": (jd_text or body or "")[:12000],
            "quick_apply": easy,
            "notes": "easy_apply_detected" if easy else "linkedin_standard_apply",
        }

    def is_valid_linkedin_job(title: str, company: str, jd_text: str) -> bool:
        t = (title or "").strip().lower()
        c = (company or "").strip().lower()
        j = (jd_text or "").strip().lower()
        junk_markers = [
            "join linkedin",
            "sign in",
            "登录",
            "加入领英",
            "create your profile",
            "forgot password",
            "agree & join",
        ]
        if any(x in t for x in junk_markers):
            return False
        if any(x in j for x in ["同意并加入", "用户协议", "cookie 政策", "forgot password"]) and len(j) < 2500:
            return False
        if t in {"linkedin", "jobs", "home"}:
            return False
        if c in {"linkedin", "unknown company"} and len(t) < 4:
            return False
        return True

    q = (query or "").strip() or "business analyst"
    loc = (location or "").strip() or "Brisbane"
    url = f"https://www.linkedin.com/jobs/search/?keywords={quote_plus(q)}&location={quote_plus(loc)}"
    jobs: List[dict] = []
    with sync_playwright() as p:
        browser = None
        context = None
        created_context = False
        if use_system_edge:
            browser = p.chromium.connect_over_cdp(cdp_url)
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            created_context = len(browser.contexts) == 0
        else:
            browser = p.chromium.launch(channel=channel, headless=headless)
            context = browser.new_context()
        page = context.new_page()
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(3200)
        page.mouse.wheel(0, 2400)
        page.wait_for_timeout(1200)
        links = page.locator("a[href*='/jobs/view/']").all()[: max(limit * 8, 80)]
        seen = set()
        for a in links:
            try:
                href = a.get_attribute("href") or ""
                title = dedupe_repeated_phrase(a.inner_text() or "")
            except Exception:
                continue
            if not href:
                continue
            href = href.split("?")[0]
            if href.startswith("/"):
                href = urljoin("https://www.linkedin.com", href)
            elif href.startswith("www."):
                href = "https://" + href
            if href in seen:
                continue
            seen.add(href)
            fallback_company = ""
            try:
                card = a.locator("xpath=ancestor::*[self::li or self::div][1]")
                if card.count() > 0:
                    fallback_company = norm_txt(
                        card.locator(
                            "h4, .base-search-card__subtitle, .job-search-card__subtitle-link, .artdeco-entity-lockup__subtitle"
                        ).first.inner_text(timeout=900)
                    )
            except Exception:
                fallback_company = ""
            parsed = {
                "title": title or "LinkedIn role",
                "company": fallback_company or "Unknown company",
                "location": loc,
                "jd_text": "",
                "notes": "linkedin_scraped_basic",
                "quick_apply": "easy apply" in (title or "").lower(),
            }
            try:
                page2 = context.new_page()
                page2.goto(href, wait_until="domcontentloaded", timeout=45000)
                page2.wait_for_timeout(1200)
                parsed = parse_linkedin_detail(page2, fallback_title=title, fallback_company=fallback_company, fallback_loc=loc)
                page2.close()
            except Exception:
                pass
            if not is_valid_linkedin_job(parsed["title"], parsed["company"], parsed["jd_text"]):
                continue
            jobs.append(
                {
                    "job_url": href,
                    "title": dedupe_repeated_phrase(parsed["title"]),
                    "company": parsed["company"],
                    "location": parsed["location"],
                    "jd_text": parsed["jd_text"],
                    "notes": parsed["notes"],
                    "quick_apply": parsed["quick_apply"],
                }
            )
            if len(jobs) >= limit:
                break
        if not use_system_edge or created_context:
            context.close()
        if not use_system_edge and browser is not None:
            browser.close()
    return jobs[:limit]


def write_jobs_csv(rows: List[dict], platform: str, append: bool = False) -> None:
    fields = ["job_id", "platform", "company", "title", "location", "job_url", "jd_text", "notes"]
    existing: List[dict] = []
    if append and JOBS_FILE.exists():
        with JOBS_FILE.open("r", encoding="utf-8-sig", newline="") as f:
            existing = list(csv.DictReader(f))
    out = existing[:]
    start_idx = len(existing) + 1
    for i, r in enumerate(rows, start=start_idx):
        out.append(
            {
                "job_id": r.get("job_id") or f"{platform}_scrape_{i:03d}",
                "platform": platform,
                "company": r.get("company", ""),
                "title": r.get("title", ""),
                "location": r.get("location", ""),
                "job_url": r.get("job_url", ""),
                "jd_text": r.get("jd_text", ""),
                "notes": r.get("notes", ""),
            }
        )
    try:
        with JOBS_FILE.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(out)
    except PermissionError:
        fallback = JOBS_FILE.with_name("jobs_latest.csv")
        with fallback.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(out)
        print(f"Warning: jobs.csv is locked. Wrote to fallback file: {fallback}")


def assist_apply_dialog(
    platform: str,
    limit: int = 20,
    headless: bool = False,
    auto_click: bool = False,
    auto_fill: bool = False,
    auto_submit: bool = False,
    channel: str = "msedge",
    use_system_edge: bool = False,
    cdp_url: str = "http://127.0.0.1:9222",
) -> None:
    if not AUDIT_XLSX_FILE.exists():
        raise FileNotFoundError(f"Audit file not found: {AUDIT_XLSX_FILE}. Run pipeline first.")

    rows = read_xlsx_rows(AUDIT_XLSX_FILE)
    pending_pairs: List[Tuple[int, dict]] = []
    for idx, row in enumerate(rows):
        if row.get("platform", "").lower() == platform.lower() and row.get("status", "") == "ready_to_apply":
            pending_pairs.append((idx, row))
    pending_pairs = pending_pairs[:limit]

    if not pending_pairs:
        print(f"No pending rows for platform={platform}")
        return

    profile = load_candidate_profile()

    def detect_submission_success(page, platform_name: str) -> bool:
        try:
            current_url = page.url.lower()
        except Exception:
            current_url = ""
        try:
            body_text = (page.inner_text("body") or "").lower()
        except Exception:
            body_text = ""
        markers = [
            "application submitted",
            "thanks for applying",
            "thank you for applying",
            "you've applied",
            "your application has been submitted",
        ]
        if any(x in body_text for x in markers):
            return True
        if platform_name.lower() == "linkedin":
            return ("/jobs/view/" in current_url and "submittedapplication=true" in current_url) or ("application sent" in body_text)
        if platform_name.lower() == "seek":
            return ("seek.com.au" in current_url and ("application" in current_url and "submitted" in current_url))
        return False

    def advance_application_steps(page, platform_name: str, max_steps: int = 5) -> List[str]:
        steps: List[str] = []
        base_selectors = [
            "button:has-text('Continue')",
            "button:has-text('Next')",
            "button:has-text('Review')",
            "button:has-text('Review application')",
            "a:has-text('Continue')",
        ]
        if platform_name.lower() == "linkedin":
            base_selectors = [
                "button:has-text('Continue to next step')",
                "button:has-text('Next')",
                "button:has-text('Review')",
                "button:has-text('Review application')",
                "button[aria-label*='Continue']",
            ] + base_selectors
        if platform_name.lower() == "seek":
            base_selectors = [
                "button:has-text('Continue')",
                "button:has-text('Save and continue')",
                "a:has-text('Continue')",
            ] + base_selectors
        for _ in range(max_steps):
            moved = click_first(page, base_selectors)
            if not moved:
                break
            steps.append("advanced_step")
            page.wait_for_timeout(900)
        return steps

    with sync_playwright() as p:
        browser = None
        context = None
        if use_system_edge:
            try:
                browser = p.chromium.connect_over_cdp(cdp_url)
            except Exception as e:
                raise RuntimeError(
                    f"Failed to connect system Edge CDP at {cdp_url}. "
                    f"Start with: msedge --remote-debugging-port=9222"
                ) from e
            context = browser.contexts[0] if browser.contexts else browser.new_context()
        else:
            profile_path = PROFILE_DIR / platform.lower()
            if not profile_path.exists():
                raise FileNotFoundError(
                    f"Missing browser profile for {platform}. Run auth first: "
                    f"py scripts/job_workflow.py auth --platform {platform}"
                )
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(profile_path),
                headless=headless,
                viewport={"width": 1440, "height": 900},
                channel=channel,
                args=["--disable-blink-features=AutomationControlled"],
            )
        page = context.new_page()
        for idx, row in pending_pairs:
            url = str(row.get("job_url", "")).strip()
            if not url:
                rows[idx]["status"] = "skipped_no_url"
                continue

            print("=" * 60)
            print(f"[Dialog] {row.get('title', '')} @ {row.get('company', '')}")
            print(f"URL: {url}")
            print(f"Resume: {row.get('resume_file', '')}")
            print(f"Cover Letter: {row.get('cover_letter_file', '')}")
            try:
                page.goto(url, wait_until="domcontentloaded")
            except PlaywrightTimeoutError:
                pass

            auto_msg = ""
            if auto_click:
                chain = click_apply_chain(page)
                if not chain:
                    auto_msg = try_auto_click_apply(page, platform)
                else:
                    auto_msg = " -> ".join(chain)
                print(f"Auto step: {auto_msg}")

            fill_stats = {"fields": 0, "uploads": 0, "coverletter": 0}
            if auto_fill:
                cover_text = read_cover_letter_text(str(row.get("cover_letter_file", "")))
                fill_stats = autofill_form(
                    page,
                    profile=profile,
                    resume_file=str(row.get("resume_file", "")),
                    cover_letter_text=cover_text,
                )
                auto_msg = f"{auto_msg}; autofill fields={fill_stats['fields']} upload={fill_stats['uploads']} cover={fill_stats['coverletter']}".strip("; ")
                print(f"Autofill: {auto_msg}")

            submitted = False
            if auto_submit:
                advance_steps = advance_application_steps(page, platform_name=platform, max_steps=5)
                if advance_steps:
                    auto_msg = f"{auto_msg}; {' -> '.join(advance_steps)}".strip("; ")
                submitted = click_first(
                    page,
                    [
                        "button:has-text('Submit')",
                        "button:has-text('Submit application')",
                        "button:has-text('Submit Application')",
                        "button:has-text('Send application')",
                        "button:has-text('Apply now')",
                        "button:has-text('Done')",
                        "input[type='submit']",
                    ],
                )
                if not submitted:
                    submitted = detect_submission_success(page, platform)
                if submitted:
                    auto_msg = f"{auto_msg}; submitted"
                    print("Auto submit: submitted")
                else:
                    auto_msg = f"{auto_msg}; submit_not_found"
                    print("Auto submit: submit button not found")
            # Safety lock: if we couldn't upload resume, never mark submitted automatically.
            if auto_fill and "upload=0" in auto_msg and submitted:
                submitted = False
                auto_msg = f"{auto_msg}; blocked_no_resume_upload"
                print("Safety lock: blocked submit because resume upload was not detected.")

            if submitted:
                decision = "submitted"
            elif auto_fill or auto_click:
                decision = "ready_for_review"
            else:
                decision = "hold"
            rows[idx]["status"] = decision
            rows[idx]["auto_step"] = auto_msg
            rows[idx]["applied_at"] = datetime.now().isoformat(timespec="seconds")
        if not use_system_edge:
            context.close()
        if browser is not None:
            browser.close()

    write_xlsx(rows, AUDIT_XLSX_FILE)
    print(f"Audit updated: {AUDIT_XLSX_FILE}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Resume tailoring + application workflow")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("init", help="Create folders + starter files")

    auth = sub.add_parser("auth", help="Open browser for manual login and save session")
    auth.add_argument("--platform", required=True, choices=["seek", "linkedin"])
    auth.add_argument("--headless", action="store_true")
    auth.add_argument("--wait-seconds", type=int, default=120)
    auth.add_argument("--channel", choices=["msedge", "chrome"], default="msedge")
    auth.add_argument("--use-system-edge", action="store_true")
    auth.add_argument("--cdp-port", type=int, default=9222)

    run = sub.add_parser("run", help="Build tailored resumes + score + report")
    run.add_argument("--use-openai", action="store_true", help="Use OPENAI_API_KEY for rewrite")

    scrape_cmd = sub.add_parser("scrape", help="Scrape jobs only (no application)")
    scrape_cmd.add_argument("--platform", required=True, choices=["seek", "linkedin"])
    scrape_cmd.add_argument("--query", default="business graduate")
    scrape_cmd.add_argument("--location", default="Brisbane QLD")
    scrape_cmd.add_argument("--limit", type=int, default=20)
    scrape_cmd.add_argument("--append", action="store_true", help="Append to existing jobs.csv")
    scrape_cmd.add_argument("--channel", choices=["msedge", "chrome"], default="msedge")
    scrape_cmd.add_argument("--headless", action="store_true")
    scrape_cmd.add_argument("--use-system-edge", action="store_true")
    scrape_cmd.add_argument("--cdp-url", default="http://127.0.0.1:9222")

    apply_cmd = sub.add_parser("apply", help="Assist manual applying using saved session")
    apply_cmd.add_argument("--platform", required=True, choices=["seek", "linkedin"])
    apply_cmd.add_argument("--limit", type=int, default=20)
    apply_cmd.add_argument("--headless", action="store_true")
    apply_cmd.add_argument("--auto-click", action="store_true", help="Try clicking Apply/Easy Apply automatically")
    apply_cmd.add_argument("--auto-fill", action="store_true", help="Auto-fill detected application fields")
    apply_cmd.add_argument("--auto-submit", action="store_true", help="Attempt final submit automatically")
    apply_cmd.add_argument("--channel", choices=["msedge", "chrome"], default="msedge")
    apply_cmd.add_argument("--use-system-edge", action="store_true")
    apply_cmd.add_argument("--cdp-url", default="http://127.0.0.1:9222")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.cmd == "init":
        init_files()
        print("Initialized project scaffolding.")
    elif args.cmd == "auth":
        login_platform(
            args.platform,
            headless=args.headless,
            wait_seconds=args.wait_seconds,
            channel=args.channel,
            use_system_edge=args.use_system_edge,
            cdp_port=args.cdp_port,
        )
    elif args.cmd == "run":
        run_pipeline(use_openai=args.use_openai)
    elif args.cmd == "scrape":
        if args.platform == "seek":
            rows = scrape_seek_jobs(
                args.query,
                args.location,
                args.limit,
                channel=args.channel,
                headless=args.headless,
                use_system_edge=args.use_system_edge,
                cdp_url=args.cdp_url,
            )
        else:
            rows = scrape_linkedin_jobs(
                args.query,
                args.location,
                args.limit,
                channel=args.channel,
                headless=args.headless,
                use_system_edge=args.use_system_edge,
                cdp_url=args.cdp_url,
            )
        write_jobs_csv(rows, platform=args.platform, append=args.append)
        print(f"Scraped {len(rows)} jobs to: {JOBS_FILE}")
    elif args.cmd == "apply":
        assist_apply_dialog(
            platform=args.platform,
            limit=args.limit,
            headless=args.headless,
            auto_click=args.auto_click,
            auto_fill=args.auto_fill,
            auto_submit=args.auto_submit,
            channel=args.channel,
            use_system_edge=args.use_system_edge,
            cdp_url=args.cdp_url,
        )


if __name__ == "__main__":
    main()
