from __future__ import annotations

import re
from typing import Dict, List, Sequence

from generation_models import JDSignals
from role_classifier import classify_role_family


NOISE_LINE_PATTERNS = [
    r"\babout us\b",
    r"\babout the company\b",
    r"\babout this opportunity\b",
    r"\babout you\b",
    r"\bwhy join\b",
    r"\bwho we are\b",
    r"\bwhat we offer\b",
    r"\bwhat's in it for you\b",
    r"\bbenefits?\b",
    r"\bperks?\b",
    r"\bwellbeing\b",
    r"\bsuperannuation\b",
    r"\bhew level\b",
    r"\bemployee assistance\b",
    r"\bwork180\b",
    r"\bhow to apply\b",
    r"\bquestions\?\b",
    r"\bapply now\b",
    r"\bclosing date\b",
    r"\binterviews will be conducted\b",
    r"\bequal opportunity employer\b",
    r"\bfast[- ]paced environment\b",
    r"\bsupportive team\b",
    r"\bsupportive environment\b",
    r"\bdynamic environment\b",
    r"\bgreat team\b",
    r"\bour values\b",
    r"\bour culture\b",
    r"\baward[- ]winning\b",
    r"\bglobal leader\b",
]

NOISE_PHRASES = {
    "key responsibilities",
    "about this opportunity",
    "about you",
    "what we offer",
    "why join us",
    "hew level",
    "superannuation",
    "supportive team",
    "fast-paced environment",
    "supportive environment",
    "great team",
}

SECTION_HINTS = {
    "responsibilities": [
        r"\bkey responsibilities\b",
        r"\bresponsibilities\b",
        r"\byou will\b",
        r"\bthis role\b",
        r"\bresponsible for\b",
    ],
    "requirements": [
        r"\brequired\b",
        r"\byou bring\b",
        r"\babout your skills\b",
        r"\bqualifications\b",
        r"\bexperience\b",
        r"\byou will need\b",
        r"\bto be eligible\b",
    ],
}

ROLE_TERM_LIBRARY: Dict[str, Dict[str, List[str]]] = {
    "accounting_reporting": {
        "skills": ["tax compliance", "financial reporting", "general ledger", "reconciliations", "month-end close", "forecasting", "compliance", "accounts payable", "accounts receivable", "budgeting"],
        "tools": ["excel", "sap", "xero", "myob", "quickbooks", "erp", "accounting systems"],
        "qualifications": ["accounting degree", "finance degree", "ca", "cpa"],
        "domain": ["tax", "audit", "financial statements", "cash flow", "variance analysis"],
    },
    "accounting_systems": {
        "skills": ["accounts payable", "accounts receivable", "invoicing", "reconciliations", "month-end close", "p&l reporting", "finance operations", "billing", "cash application"],
        "tools": ["excel", "sap", "xero", "myob", "quickbooks", "erp", "accounting systems"],
        "qualifications": ["accounting degree", "finance degree", "ca", "cpa"],
        "domain": ["finance systems", "systems implementation", "internal finance", "profit and loss", "accounts operations"],
    },
    "data_analyst": {
        "skills": ["data analysis", "reporting", "dashboarding", "trend analysis", "data validation", "data quality", "data modelling", "visualisation", "stakeholder reporting"],
        "tools": ["power bi", "excel", "sql", "tableau", "sap business objects", "python"],
        "qualifications": ["analytics degree", "data science degree", "statistics degree", "mathematics degree"],
        "domain": ["dashboards", "workflows", "assessment data", "live reporting", "decision-making"],
    },
    "business_analyst": {
        "skills": ["business process improvement", "requirements gathering", "stakeholder support", "workflow documentation", "operational reporting", "process analysis", "problem solving", "documentation", "project support"],
        "tools": ["excel", "power bi", "visio", "jira", "confluence", "crm", "erp"],
        "qualifications": ["business degree", "commerce degree", "information systems degree"],
        "domain": ["business processes", "workflows", "process improvement", "stakeholders", "operational support"],
    },
    "business_graduate": {
        "skills": ["reporting", "coordination", "customer support", "commercial exposure", "administration", "adaptability", "communication", "problem solving"],
        "tools": ["excel", "crm", "power bi"],
        "qualifications": ["bachelor degree", "commerce degree", "business degree"],
        "domain": ["rotations", "operations", "customer journey", "cross-functional support", "commercial"],
    },
    "consulting": {
        "skills": ["stakeholder management", "analysis", "problem solving", "research", "recommendations"],
        "tools": ["excel", "power bi"],
        "qualifications": ["business degree", "commerce degree"],
        "domain": ["advisory", "strategy", "commercial"],
    },
    "sales_bd": {
        "skills": ["client communication", "revenue growth", "pipeline management", "relationship management", "negotiation"],
        "tools": ["crm", "excel"],
        "qualifications": ["business degree", "commerce degree"],
        "domain": ["customers", "sales", "commercial"],
    },
    "operations": {
        "skills": ["process support", "coordination", "documentation", "reporting", "service delivery"],
        "tools": ["excel", "crm", "erp"],
        "qualifications": ["business degree"],
        "domain": ["operations", "service", "process"],
    },
}

ROLE_TERM_LIBRARY["accounting"] = ROLE_TERM_LIBRARY["accounting_reporting"]


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _norm_key(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def _split_lines(jd_text: str) -> List[str]:
    lines: List[str] = []
    for chunk in re.split(r"[\r\n]+", str(jd_text or "")):
        line = _normalize(chunk)
        if line:
            lines.append(line)
    return lines


def _is_noise_line(line: str) -> bool:
    low = line.lower()
    if any(re.search(pattern, low) for pattern in NOISE_LINE_PATTERNS):
        return True
    if len(low.split()) <= 2:
        return True
    if re.fullmatch(r"[\W\d ]+", low):
        return True
    if line.endswith(":") and _norm_key(line[:-1]) in NOISE_PHRASES:
        return True
    return False


def _is_high_signal_line(line: str) -> bool:
    low = line.lower()
    if _is_noise_line(line):
        return False
    if len(low.split()) < 4:
        return False
    if re.search(r"\b(benefits?|salary|superannuation|culture|wellbeing|fitness|discounts|employee assistance)\b", low):
        return False
    return bool(
        re.search(
            r"\b(prepare|develop|analyse|analyze|support|maintain|coordinate|deliver|report|document|identify|assist|manage|build|collate|clean|model|present|partner|work with|experience in|exposure to|degree|qualification|certification|power bi|excel|sql|tableau|python|stakeholder|workflow|process)\b",
            low,
        )
    )


def clean_jd_text(jd_text: str) -> str:
    cleaned = [line for line in _split_lines(jd_text) if _is_high_signal_line(line)]
    return "\n".join(cleaned)


def _extract_sentence_phrases(line: str) -> List[str]:
    text = _norm_key(line)
    anchors = [
        "power bi",
        "sql",
        "excel",
        "tableau",
        "python",
        "sap business objects",
        "business processes",
        "business process models",
        "requirements gathering",
        "stakeholder communication",
        "stakeholder engagement",
        "stakeholder support",
        "workflow documentation",
        "process improvement",
        "operational reporting",
        "live reporting",
        "assessment data",
        "data quality",
        "trend analysis",
        "project support",
        "reporting",
        "dashboards",
        "documentation",
        "coordination",
        "commercial",
        "operations",
        "customer support",
    ]
    output: List[str] = []
    for anchor in anchors:
        if anchor in text:
            output.append(anchor)
    return output


def _collect_role_terms(cleaned_text: str, role_family: str, bucket: str) -> List[str]:
    low = cleaned_text.lower()
    terms: List[str] = []
    for term in ROLE_TERM_LIBRARY.get(role_family, {}).get(bucket, []):
        if term in low:
            terms.append(term)
    return terms


def _extract_section_lines(cleaned_text: str, section_name: str, limit: int = 8) -> List[str]:
    patterns = SECTION_HINTS.get(section_name, [])
    lines = cleaned_text.splitlines()
    selected: List[str] = []
    for line in lines:
        low = line.lower()
        if section_name == "responsibilities":
            if re.search(r"\b(prepare|develop|analyse|analyze|support|maintain|coordinate|deliver|report|document|identify|assist|manage|build|collate|clean|model|present|partner)\b", low):
                selected.append(line)
                continue
        if section_name == "requirements":
            if re.search(r"\b(experience in|exposure to|degree|qualification|certification|ability to|attention to detail|communication skills|stakeholder)\b", low):
                selected.append(line)
                continue
        if any(re.search(pattern, low) for pattern in patterns):
            continue
    deduped: List[str] = []
    seen = set()
    for line in selected:
        key = _norm_key(line)
        if key and key not in seen:
            seen.add(key)
            deduped.append(line)
        if len(deduped) >= limit:
            break
    return deduped


def _rank_keywords(cleaned_text: str, role_family: str, limit: int = 15) -> List[str]:
    candidates = (
        _collect_role_terms(cleaned_text, role_family, "skills")
        + _collect_role_terms(cleaned_text, role_family, "tools")
        + _collect_role_terms(cleaned_text, role_family, "qualifications")
        + _collect_role_terms(cleaned_text, role_family, "domain")
    )
    for line in cleaned_text.splitlines():
        candidates.extend(_extract_sentence_phrases(line))

    ranked: List[str] = []
    seen = set()
    allowed_single_terms = {"excel", "sql", "python", "reporting", "stakeholder", "compliance", "documentation", "dashboards", "commercial", "operations"}
    for phrase in candidates:
        key = _norm_key(phrase)
        if not key or key in seen:
            continue
        if key in NOISE_PHRASES:
            continue
        if len(key.split()) == 1 and key not in allowed_single_terms:
            continue
        if len(key.split()) > 4:
            continue
        if len(key) < 5:
            continue
        if re.search(r"\b(about|opportunity|benefits|culture|salary|superannuation|hew|level|team|environment|offer)\b", key):
            continue
        if re.search(r"\b(and|or|the|this|that|your|our|their)\b", key) and len(key.split()) <= 2:
            continue
        if re.search(r"^[a-z] [a-z]", key):
            continue
        ranked.append(key)
        seen.add(key)
        if len(ranked) >= limit:
            break
    return ranked


def extract_core_skills(cleaned_text: str, role_family: str, keywords: Sequence[str]) -> List[str]:
    seeded = _collect_role_terms(cleaned_text, role_family, "skills")
    for keyword in keywords:
        if any(term in keyword for term in ["report", "analysis", "dashboard", "stakeholder", "process", "workflow", "documentation", "communication", "compliance", "forecast"]):
            seeded.append(keyword)
    out = []
    seen = set()
    for item in seeded:
        key = _norm_key(item)
        if key and key not in seen:
            seen.add(key)
            out.append(key)
    return out[:10]


def extract_tools(cleaned_text: str, role_family: str) -> List[str]:
    out = _collect_role_terms(cleaned_text, role_family, "tools")
    seen = set()
    deduped: List[str] = []
    for tool in out:
        key = _norm_key(tool)
        if key and key not in seen:
            seen.add(key)
            deduped.append(key)
    return deduped[:8]


def extract_qualifications(cleaned_text: str, role_family: str) -> List[str]:
    out = _collect_role_terms(cleaned_text, role_family, "qualifications")
    lines = cleaned_text.splitlines()
    for line in lines:
        low = line.lower()
        if re.search(r"\b(degree|qualification|certification|cpa|ca|bachelor|master|commerce|business|analytics|accounting|finance)\b", low):
            out.extend(_extract_sentence_phrases(line))
    deduped: List[str] = []
    seen = set()
    for item in out:
        key = _norm_key(item)
        if key and key not in seen:
            seen.add(key)
            deduped.append(key)
    return deduped[:8]


def extract_domain_terms(cleaned_text: str, role_family: str) -> List[str]:
    terms = _collect_role_terms(cleaned_text, role_family, "domain")
    deduped: List[str] = []
    seen = set()
    for item in terms:
        key = _norm_key(item)
        if key and key not in seen:
            seen.add(key)
            deduped.append(key)
    return deduped[:6]


def parse_jd(job_title: str, jd_text: str) -> JDSignals:
    """Parse a JD into clean, role-relevant ATS signals.

    Only responsibilities, requirements, tools, qualifications, and role-relevant
    domain terms survive. Branding, benefits, HR policy, and sentence fragments
    are aggressively removed upstream so generation and scoring share cleaner inputs.
    """
    cleaned = clean_jd_text(jd_text)
    role_family = classify_role_family(job_title, cleaned)
    keywords = _rank_keywords(cleaned, role_family, limit=16)
    responsibilities = _extract_section_lines(cleaned, "responsibilities", limit=10)
    requirements = _extract_section_lines(cleaned, "requirements", limit=8)
    tools = extract_tools(cleaned, role_family)
    qualifications = extract_qualifications(cleaned, role_family)
    domain_terms = extract_domain_terms(cleaned, role_family)
    core_skills = extract_core_skills(cleaned, role_family, keywords)

    return JDSignals(
        role_family=role_family,
        target_title=job_title,
        cleaned_text=cleaned,
        keywords=keywords,
        core_skills=core_skills or keywords[:8],
        responsibilities=responsibilities or requirements,
        tools=tools,
        qualifications=qualifications,
        domain_terms=domain_terms,
    )
