from __future__ import annotations

import re
from typing import Dict, List, Tuple

from generation_models import ATSResult, JDSignals, ResumeAnalysis


SECTION_HEADERS = {"SUMMARY", "SKILLS", "EXPERIENCE", "EDUCATION", "AVAILABILITY"}

GENERIC_SUMMARY_PATTERNS = [
    r"\bresults[- ]driven\b",
    r"\bhighly motivated\b",
    r"\bdynamic professional\b",
    r"\bdetail-oriented professional\b",
    r"\bcandidate aligned to\b",
    r"\bprofessional aligned to\b",
    r"\bpractical experience in key responsibilities\b",
]

NOISE_PATTERNS = [
    r"\baward-winning\b",
    r"\bglobal leader\b",
    r"\bfast-paced environment\b",
    r"\bgreat team\b",
    r"\bculture\b",
    r"\bbenefits?\b",
    r"\bperks?\b",
    r"\bemployee assistance\b",
    r"\bsuperannuation\b",
    r"\bhew level\b",
]

LOW_VALUE_KEYWORD_PATTERNS = [
    r"\bsupportive environment\b",
    r"\bfast[- ]paced\b",
    r"\bdynamic\b",
    r"\bgreat team\b",
    r"\btraining\b",
    r"\bgym\b",
    r"\bculture\b",
    r"\bbenefits?\b",
    r"\baward[- ]winning\b",
    r"\bglobal leader\b",
    r"\brecently completed\b",
    r"\babout this opportunity\b",
    r"\bkey responsibilities\b",
    r"\bhew level\b",
    r"\bsuperannuation\b",
]

ROLE_WEIGHT_RULES = {
    "accounting": {
        "primary": ["tax", "taxation", "compliance", "financial reporting", "general ledger", "reconciliations", "month end", "excel", "accounting systems", "clients"],
        "secondary": ["forecasting", "variance analysis", "communication", "attention to detail", "finance", "reporting"],
    },
    "accounting_reporting": {
        "primary": ["tax", "taxation", "compliance", "financial reporting", "financial statements", "reconciliations", "month end", "excel", "forecasting"],
        "secondary": ["reporting", "accuracy", "finance", "variance analysis", "cash flow"],
    },
    "accounting_systems": {
        "primary": ["accounts payable", "accounts receivable", "invoicing", "billing", "reconciliations", "month end", "p l reporting", "profit and loss", "xero", "myob", "quickbooks", "erp", "sap", "accounting systems"],
        "secondary": ["finance operations", "reporting", "accuracy", "cash flow", "excel"],
    },
    "data_analyst": {
        "primary": ["sql", "python", "power bi", "tableau", "dashboard", "data analysis"],
        "secondary": ["forecasting", "stakeholder", "reporting", "visualisation", "data quality", "trend analysis"],
    },
    "business_analyst": {
        "primary": ["business process improvement", "requirements gathering", "workflow documentation", "stakeholder support", "operational reporting", "process analysis"],
        "secondary": ["documentation", "coordination", "problem solving", "reporting", "stakeholder", "workflow"],
    },
    "business_graduate": {
        "primary": ["reporting", "coordination", "customer support", "commercial", "operations"],
        "secondary": ["adaptability", "communication", "cross-functional", "support", "execution"],
    },
    "consulting": {
        "primary": ["stakeholder", "problem solving", "analysis", "communication"],
        "secondary": ["workshop", "reporting", "recommendations"],
    },
    "sales_bd": {
        "primary": ["client", "revenue", "negotiation", "growth", "pipeline"],
        "secondary": ["communication", "relationship", "retention"],
    },
    "operations": {
        "primary": ["process", "efficiency", "coordination", "operations"],
        "secondary": ["reporting", "documentation", "service"],
    },
}


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def _split_sections(resume_text: str) -> Dict[str, List[str]]:
    sections: Dict[str, List[str]] = {key: [] for key in SECTION_HEADERS}
    current = ""
    for raw in str(resume_text or "").splitlines():
        line = raw.strip()
        if not line:
            continue
        upper = line.upper()
        if upper in SECTION_HEADERS:
            current = upper
            continue
        if current:
            sections[current].append(line)
    return sections


def _keyword_tier(keyword: str, role_family: str) -> int:
    key = _norm(keyword)
    role_rules = ROLE_WEIGHT_RULES.get(role_family, {})
    primary = role_rules.get("primary", [])
    secondary = role_rules.get("secondary", [])
    if any(term in key for term in primary):
        return 1
    if any(term in key for term in secondary):
        return 2
    if any(term in key for term in ["report", "analysis", "process", "communication", "detail", "tool"]):
        return 2
    return 3


def _is_low_value_keyword(keyword: str) -> bool:
    key = keyword.lower().strip()
    if not key or len(key) < 3:
        return True
    if len(key.split()) > 6:
        return True
    if re.search(r"[^a-z0-9&/ +#-]", key):
        return True
    if re.search(r"\b(if|you|we|our|your|they)\b", key) and len(key.split()) <= 2:
        return True
    if re.search(r"\b(if|you|we|our|your|they|recently|completed)\b", key) and len(key.split()) >= 4:
        return True
    return any(re.search(pattern, key) for pattern in LOW_VALUE_KEYWORD_PATTERNS)


def _keyword_weight(keyword: str, role_family: str) -> float:
    tier = _keyword_tier(keyword, role_family)
    base = {1: 5.0, 2: 3.0, 3: 1.0}[tier]
    key = _norm(keyword)
    role_rules = ROLE_WEIGHT_RULES.get(role_family, {})
    if any(term in key for term in role_rules.get("primary", [])):
        base *= 1.2
    return base


def _keyword_context_score(keyword: str, sections: Dict[str, List[str]]) -> Tuple[float, Dict[str, bool]]:
    key = _norm(keyword)
    found = False
    in_summary = False
    in_skills = False
    in_experience = False
    quantified = False
    flat_skill_dump = False

    for section_name, lines in sections.items():
        for line in lines:
            normalized = _norm(line)
            if key not in normalized:
                continue
            found = True
            if section_name == "SUMMARY":
                in_summary = True
            elif section_name == "SKILLS":
                in_skills = True
                if line.count(",") >= 4 and not re.search(r"\b(use|prepare|manage|apply|support|communicate|analyse|analyze)\b", line.lower()):
                    flat_skill_dump = True
            elif section_name == "EXPERIENCE":
                in_experience = True
                if re.search(r"(\d+|%|\$|million|improved|reduced|generated|maintained|accuracy)", line.lower()):
                    quantified = True

    if not found:
        return 0.0, {
            "found": False,
            "summary": False,
            "skills": False,
            "experience": False,
            "quantified": False,
            "flat_skill_dump": False,
        }

    score = 0.0
    if in_summary:
        score += 0.6
    if in_skills:
        score += 0.5
    if in_experience:
        score += 1.4
    if quantified:
        score += 0.7
    if flat_skill_dump:
        score -= 0.35
    return max(score, 0.2), {
        "found": True,
        "summary": in_summary,
        "skills": in_skills,
        "experience": in_experience,
        "quantified": quantified,
        "flat_skill_dump": flat_skill_dump,
    }


def _core_keyword_coverage(jd_signals: JDSignals, sections: Dict[str, List[str]]) -> Tuple[float, List[str], List[str], List[str], List[str], float]:
    matched_keywords: List[str] = []
    missing_keywords: List[str] = []
    matched_tier1: List[str] = []
    missing_tier1: List[str] = []
    total_possible = 0.0
    total_scored = 0.0

    signal_pool = list(dict.fromkeys(jd_signals.keywords + jd_signals.core_skills + jd_signals.tools + jd_signals.qualifications + jd_signals.domain_terms))
    keywords = [keyword for keyword in signal_pool if not _is_low_value_keyword(keyword)]

    for keyword in keywords:
        weight = _keyword_weight(keyword, jd_signals.role_family)
        total_possible += weight * 2.1
        ctx_score, flags = _keyword_context_score(keyword, sections)
        if not flags["found"]:
            missing_keywords.append(keyword)
            if _keyword_tier(keyword, jd_signals.role_family) == 1:
                missing_tier1.append(keyword)
            continue
        total_scored += weight * ctx_score
        matched_keywords.append(keyword)
        if _keyword_tier(keyword, jd_signals.role_family) == 1:
            matched_tier1.append(keyword)

    coverage_score = round(100 * total_scored / max(total_possible, 1.0), 1)
    return coverage_score, matched_keywords[:20], missing_keywords[:15], matched_tier1[:12], missing_tier1[:12], total_scored


def _experience_relevance_score(analysis: ResumeAnalysis, sections: Dict[str, List[str]]) -> float:
    experience_lines = sections.get("EXPERIENCE", [])
    experience_text = _norm(" ".join(experience_lines))
    if not experience_text:
        return 0.0
    preserve_hits = sum(1 for role in analysis.keep if _norm(role.header) in experience_text)
    boost_hits = sum(1 for role in analysis.boost if _norm(role.header) in experience_text)
    total_targets = max(len(analysis.keep) + len(analysis.boost), 1)
    return round(100 * ((preserve_hits * 1.0) + (boost_hits * 0.7)) / total_targets, 1)


def _quantified_impact_score(sections: Dict[str, List[str]], jd_signals: JDSignals) -> float:
    experience_lines = sections.get("EXPERIENCE", [])
    if not experience_lines:
        return 0.0
    role_terms = [_norm(x) for x in jd_signals.keywords[:10]]
    hits = 0
    total = 0
    for line in experience_lines:
        if not line.startswith("-"):
            continue
        total += 1
        low = line.lower()
        has_metric = bool(re.search(r"(\d+|%|\$|million|accuracy|roi)", low))
        has_role_term = any(term and term in _norm(low) for term in role_terms)
        if has_metric and has_role_term:
            hits += 1
    return round(100 * hits / max(total, 1), 1)


def _role_title_alignment_score(resume_text: str, jd_signals: JDSignals) -> float:
    resume_norm = _norm(resume_text)
    title_norm = _norm(jd_signals.target_title)
    if not title_norm:
        return 0.0
    if title_norm in resume_norm:
        return 100.0
    words = [w for w in title_norm.split() if len(w) > 3]
    hits = sum(1 for word in words if word in resume_norm)
    return round(100 * hits / max(len(words), 1), 1)


def _structural_completeness_score(sections: Dict[str, List[str]]) -> float:
    required = ["SUMMARY", "SKILLS", "EXPERIENCE", "EDUCATION", "AVAILABILITY"]
    present = sum(1 for section in required if sections.get(section))
    return round(100 * present / len(required), 1)


def _noise_penalty(resume_text: str) -> float:
    low = resume_text.lower()
    penalty = 0.0
    for pattern in NOISE_PATTERNS:
        if re.search(pattern, low):
            penalty += 3.0
    return min(penalty, 15.0)


def _genericity_penalty(sections: Dict[str, List[str]]) -> float:
    penalty = 0.0
    summary_text = " ".join(sections.get("SUMMARY", []))
    if any(re.search(pattern, summary_text.lower()) for pattern in GENERIC_SUMMARY_PATTERNS):
        penalty += 8.0
    skills_lines = sections.get("SKILLS", [])
    for line in skills_lines:
        if line.count(",") >= 6:
            penalty += 2.5
    return min(penalty, 15.0)


def _weak_sections(
    coverage_score: float,
    contextual_quality: float,
    experience_relevance: float,
    quantified_impact: float,
    role_title_alignment: float,
    structural_completeness: float,
) -> List[str]:
    weak: List[str] = []
    if coverage_score < 75 or role_title_alignment < 80:
        weak.append("summary")
    if contextual_quality < 70:
        weak.append("skills")
    if experience_relevance < 75 or quantified_impact < 45:
        weak.append("experience")
    if structural_completeness < 100:
        weak.append("structure")
    return weak


def _suggested_refinements(missing_tier1: List[str], weak_sections: List[str]) -> List[str]:
    suggestions: List[str] = []
    if missing_tier1:
        suggestions.append(f"Add contextual evidence for tier-1 requirements: {', '.join(missing_tier1[:3])}.")
    if "summary" in weak_sections:
        suggestions.append("Refine the summary so the target role, strongest verified achievements, and core tools appear naturally in full sentences.")
    if "skills" in weak_sections:
        suggestions.append("Tighten skills into fewer role-relevant categories and replace flat lists with capability-oriented wording.")
    if "experience" in weak_sections:
        suggestions.append("Strengthen selected experience bullets by tying role-relevant skills to measurable business outcomes.")
    if "structure" in weak_sections:
        suggestions.append("Restore or strengthen required ATS-safe sections: SUMMARY, SKILLS, EXPERIENCE, EDUCATION, AVAILABILITY.")
    return suggestions[:3]


def score_resume(resume_text: str, jd_signals: JDSignals, analysis: ResumeAnalysis) -> ATSResult:
    """Score a generated resume as a generation quality gate.

    The scorer is designed to control regeneration, not just check output after
    the fact. It rewards role-relevant, contextual evidence and penalizes noise,
    genericity, and shallow keyword dumping.
    """
    sections = _split_sections(resume_text)

    coverage_score, matched_keywords, missing_keywords, matched_tier1, missing_tier1, matched_points = _core_keyword_coverage(
        jd_signals, sections
    )
    experience_relevance = _experience_relevance_score(analysis, sections)
    quantified_impact = _quantified_impact_score(sections, jd_signals)
    role_title_alignment = _role_title_alignment_score(resume_text, jd_signals)
    structural_completeness = _structural_completeness_score(sections)
    contextual_quality = round(min(((coverage_score * 0.45) + (experience_relevance * 0.35) + (quantified_impact * 0.20)), 100.0), 1)
    noise_penalty = _noise_penalty(resume_text)
    genericity_penalty = _genericity_penalty(sections)

    raw_score = (
        coverage_score * 0.28
        + contextual_quality * 0.22
        + experience_relevance * 0.18
        + quantified_impact * 0.12
        + role_title_alignment * 0.08
        + structural_completeness * 0.12
    )
    final_score = round(max(min(raw_score - noise_penalty - genericity_penalty, 100.0), 0.0), 1)

    weak_sections = _weak_sections(
        coverage_score=coverage_score,
        contextual_quality=contextual_quality,
        experience_relevance=experience_relevance,
        quantified_impact=quantified_impact,
        role_title_alignment=role_title_alignment,
        structural_completeness=structural_completeness,
    )
    improvements = _suggested_refinements(missing_tier1, weak_sections)

    return ATSResult(
        score=final_score,
        matched_keywords=matched_keywords,
        missing_keywords=missing_keywords,
        matched_tier1_keywords=matched_tier1,
        missing_tier1_keywords=missing_tier1,
        weak_sections=weak_sections,
        improvements=improvements,
        debug={
            "coverage_score": coverage_score,
            "contextual_quality": contextual_quality,
            "experience_relevance": experience_relevance,
            "quantified_impact": quantified_impact,
            "role_title_alignment": role_title_alignment,
            "structural_completeness": structural_completeness,
            "noise_penalty": noise_penalty,
            "genericity_penalty": genericity_penalty,
            "matched_points": round(matched_points, 2),
        },
    )
