from __future__ import annotations

import re
from typing import Dict, List

from generation_models import JDSignals, ResumeAnalysis, ResumeRole
from role_classifier import define_target_profile, is_accounting_family


def parse_base_resume_text(text: str) -> Dict[str, object]:
    """Parse the user's base resume into reusable structural sections."""
    data: Dict[str, object] = {
        "name": "",
        "title": "",
        "contact": "",
        "links": "",
        "summary": [],
        "skills": [],
        "experience": [],
        "education": [],
        "projects": [],
        "availability": [],
    }
    lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
    if not lines:
        return data

    data["name"] = lines[0] if len(lines) > 0 else ""
    data["title"] = lines[1] if len(lines) > 1 else ""
    data["contact"] = lines[2] if len(lines) > 2 else ""
    data["links"] = lines[3] if len(lines) > 3 and "linkedin" in lines[3].lower() else ""

    section = None
    current_role = None
    for line in lines[4:]:
        upper = line.upper()
        if upper in {"SUMMARY", "SKILLS", "EXPERIENCE", "EDUCATION", "PROJECTS", "AVAILABILITY"}:
            section = upper.lower()
            current_role = None
            continue

        if section == "summary":
            data["summary"].append(line)
        elif section == "skills":
            data["skills"].append(line)
        elif section == "experience":
            if "|" in line and not line.startswith("-"):
                current_role = {"header": line, "bullets": []}
                data["experience"].append(current_role)
            elif current_role is not None:
                bullet = line[2:].strip() if line.startswith("- ") else line
                current_role["bullets"].append(bullet)
        elif section == "education":
            data["education"].append(line)
        elif section == "projects":
            data["projects"].append(line)
        elif section == "availability":
            data["availability"].append(line)
    return data


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def _role_relevance_score(header: str, bullets: List[str], jd_signals: JDSignals) -> float:
    combined = _norm(header + " " + " ".join(bullets))
    score = 0.0
    for keyword in jd_signals.keywords:
        key = _norm(keyword)
        if key and key in combined:
            score += 2.0
    for keyword in jd_signals.core_skills:
        key = _norm(keyword)
        if key and key in combined:
            score += 3.0
    for keyword in jd_signals.tools + jd_signals.domain_terms + jd_signals.qualifications:
        key = _norm(keyword)
        if key and key in combined:
            score += 1.8
    if is_accounting_family(jd_signals.role_family) and any(x in combined for x in ["account", "tax", "financial", "ledger", "reconcil", "report"]):
        score += 4.0
    if jd_signals.role_family == "accounting_systems" and any(x in combined for x in ["accounts payable", "accounts receivable", "invoice", "billing", "erp", "sap", "quickbooks", "xero", "myob", "p l", "profit and loss"]):
        score += 4.4
    if jd_signals.role_family == "data_analyst" and any(x in combined for x in ["dashboard", "power bi", "sql", "analysis", "report", "trend", "data"]):
        score += 4.0
    if jd_signals.role_family == "business_analyst" and any(x in combined for x in ["process", "workflow", "stakeholder", "report", "documentation", "operations", "coordination"]):
        score += 4.0
    if jd_signals.role_family == "business_graduate" and any(x in combined for x in ["support", "customer", "operations", "report", "team", "coordination", "commercial"]):
        score += 3.4
    if re.search(r"(\d+|%|\$|million|kpi|roi|accuracy)", " ".join(bullets).lower()):
        score += 2.0
    return score


def _role_evidence(header: str, bullets: List[str], jd_signals: JDSignals) -> List[str]:
    combined = _norm(header + " " + " ".join(bullets))
    out: List[str] = []
    for keyword in jd_signals.keywords:
        key = _norm(keyword)
        if key and key in combined:
            out.append(keyword)
    return out[:6]


def analyze_resume(base_resume_text: str, jd_signals: JDSignals) -> ResumeAnalysis:
    """Classify resume content into KEEP / BOOST / DOWNPLAY buckets.

    The analyzer never deletes roles. It only decides what should be emphasized
    or softened during generation.
    """
    parsed = parse_base_resume_text(base_resume_text)
    keep: List[ResumeRole] = []
    boost: List[ResumeRole] = []
    downplay: List[ResumeRole] = []
    preserved_headers: List[str] = []

    for role in parsed.get("experience", []):
        header = str(role.get("header", "")).strip()
        bullets = [str(x).strip() for x in role.get("bullets", []) if str(x).strip()]
        if not header:
            continue
        preserved_headers.append(header)
        score = _role_relevance_score(header, bullets, jd_signals)
        evidence = _role_evidence(header, bullets, jd_signals)
        if score >= 6:
            keep.append(ResumeRole(header=header, bullets=bullets, classification="KEEP", relevance_score=score, evidence=evidence))
        elif score >= 3:
            boost.append(ResumeRole(header=header, bullets=bullets, classification="BOOST", relevance_score=score, evidence=evidence))
        else:
            downplay.append(ResumeRole(header=header, bullets=bullets, classification="DOWNPLAY", relevance_score=score, evidence=evidence))

    resume_text_norm = _norm(base_resume_text)
    missing = []
    for keyword in jd_signals.keywords:
        key = _norm(keyword)
        if key and key not in resume_text_norm:
            missing.append(keyword)

    return ResumeAnalysis(
        keep=keep,
        boost=boost,
        downplay=downplay,
        missing_keywords=missing[:10],
        preserved_headers=preserved_headers,
        target_profile=define_target_profile(jd_signals.role_family, jd_signals.keywords),
    )
