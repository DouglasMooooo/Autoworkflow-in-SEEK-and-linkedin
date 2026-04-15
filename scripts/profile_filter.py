from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Sequence, Tuple

from generation_models import JDSignals
from role_classifier import is_accounting_family


ROLE_FACT_BUDGETS = {
    "accounting": 10,
    "accounting_reporting": 10,
    "accounting_systems": 11,
    "data_analyst": 12,
    "business_analyst": 12,
    "business_graduate": 10,
    "consulting": 11,
    "sales_bd": 11,
    "operations": 11,
}

ROLE_PRIORITIES = {
    "accounting": {
        "boost": [
            "accounting",
            "finance",
            "tax",
            "taxation",
            "compliance",
            "reconciliation",
            "general ledger",
            "reporting",
            "audit",
            "excel",
            "erp",
            "xero",
            "myob",
            "accounts payable",
            "accounts receivable",
        ],
        "downrank": [
            "ai",
            "machine learning",
            "llm",
            "prompt",
            "automation project",
            "computer vision",
            "agent",
            "stable diffusion",
            "deep learning",
            "salary",
            "expected salary",
        ],
    },
    "accounting_reporting": {
        "boost": [
            "accounting",
            "tax",
            "taxation",
            "compliance",
            "financial reporting",
            "reconciliation",
            "month end",
            "forecast",
            "excel",
            "financial statements",
        ],
        "downrank": [
            "ai",
            "machine learning",
            "llm",
            "prompt",
            "automation project",
            "salary",
            "expected salary",
        ],
    },
    "accounting_systems": {
        "boost": [
            "accounts payable",
            "accounts receivable",
            "invoicing",
            "billing",
            "reconciliation",
            "month end",
            "p l reporting",
            "profit and loss",
            "xero",
            "myob",
            "quickbooks",
            "sap",
            "erp",
            "accounting systems",
            "finance operations",
        ],
        "downrank": [
            "ai",
            "machine learning",
            "llm",
            "prompt",
            "salary",
            "expected salary",
        ],
    },
    "data_analyst": {
        "boost": [
            "sql",
            "python",
            "dashboard",
            "power bi",
            "tableau",
            "forecast",
            "analysis",
            "visualisation",
            "stakeholder",
            "automation",
        ],
        "downrank": [
            "payroll",
            "bookkeeping",
            "tax return",
            "salary",
            "expected salary",
        ],
    },
    "business_analyst": {
        "boost": [
            "business analyst",
            "process",
            "workflow",
            "stakeholder",
            "reporting",
            "documentation",
            "requirements",
            "coordination",
            "commercial",
            "power bi",
            "excel",
        ],
        "downrank": [
            "salary",
            "expected salary",
        ],
    },
    "business_graduate": {
        "boost": [
            "graduate",
            "reporting",
            "coordination",
            "customer",
            "operations",
            "commercial",
            "adaptability",
            "support",
            "excel",
        ],
        "downrank": [
            "salary",
            "expected salary",
        ],
    },
    "consulting": {
        "boost": [
            "stakeholder",
            "analysis",
            "recommendation",
            "problem solving",
            "commercial",
            "presentation",
            "client",
            "workshop",
            "strategy",
        ],
        "downrank": [
            "salary",
            "expected salary",
        ],
    },
    "sales_bd": {
        "boost": [
            "client",
            "revenue",
            "growth",
            "pipeline",
            "negotiation",
            "partnership",
            "retention",
            "commercial",
            "relationship",
        ],
        "downrank": [
            "research publication",
            "salary",
            "expected salary",
        ],
    },
    "operations": {
        "boost": [
            "process",
            "operations",
            "coordination",
            "service",
            "reporting",
            "documentation",
            "efficiency",
            "cross-functional",
            "stakeholder",
        ],
        "downrank": [
            "salary",
            "expected salary",
        ],
    },
}

GENERIC_BRANDING_PATTERNS = [
    r"\bresults[- ]driven\b",
    r"\bpassionate professional\b",
    r"\bdynamic professional\b",
    r"\bvisionary\b",
    r"\binnovative leader\b",
    r"\bglobal mindset\b",
    r"\bpersonal brand\b",
]

QUALIFICATION_PATTERNS = [
    r"\bdegree\b",
    r"\bmajor\b",
    r"\bgraduate\b",
    r"\bgraduation\b",
    r"\bcpa\b",
    r"\bca\b",
    r"\bcertification\b",
    r"\bdiploma\b",
]

AVAILABILITY_PATTERNS = [
    r"\bavailable\b",
    r"\bstart\b",
    r"\bnotice\b",
    r"\bvisa\b",
    r"\bwork authorization\b",
    r"\bwork rights\b",
    r"\bfull-time\b",
    r"\bpart-time\b",
]

IMPACT_PATTERNS = [
    r"\$[\d,.]+",
    r"\b\d+(\.\d+)?%\b",
    r"\b\d+(\.\d+)?\s*(k|m|million|billion)\b",
    r"\bimprov",
    r"\breduc",
    r"\bincreas",
    r"\bsaved\b",
    r"\bgenerated\b",
    r"\bdelivered\b",
    r"\bidentified\b",
]

ROLE_RELEVANCE_PATTERNS = {
    "accounting": [
        r"\btax\b",
        r"\bfinancial\b",
        r"\baccount",
        r"\baudit\b",
        r"\breconcil",
        r"\bmonth end\b",
        r"\bcompliance\b",
        r"\bbudget\b",
        r"\bforecast\b",
    ],
    "accounting_reporting": [
        r"\btax\b",
        r"\bfinancial\b",
        r"\breconcil",
        r"\bmonth end\b",
        r"\bcompliance\b",
        r"\bforecast\b",
        r"\bfinancial statements\b",
    ],
    "accounting_systems": [
        r"\baccounts payable\b",
        r"\baccounts receivable\b",
        r"\binvoic",
        r"\bbilling\b",
        r"\breconcil",
        r"\bmonth end\b",
        r"\bp&l\b",
        r"\bprofit and loss\b",
        r"\bxero\b",
        r"\bmyob\b",
        r"\bquickbooks\b",
        r"\berp\b",
        r"\bsap\b",
    ],
    "data_analyst": [
        r"\bsql\b",
        r"\bpython\b",
        r"\bdashboard\b",
        r"\banalysis\b",
        r"\bpower bi\b",
        r"\btableau\b",
        r"\bvisual",
        r"\bdata\b",
    ],
    "business_analyst": [
        r"\bprocess\b",
        r"\bworkflow\b",
        r"\bstakeholder\b",
        r"\bdocument",
        r"\breport",
        r"\bcoordination\b",
        r"\boperations\b",
    ],
    "business_graduate": [
        r"\bgraduate\b",
        r"\boperations\b",
        r"\bcustomer\b",
        r"\bcommercial\b",
        r"\breport",
        r"\bcoordination\b",
        r"\bsupport\b",
    ],
    "consulting": [
        r"\bstakeholder\b",
        r"\banalysis\b",
        r"\brecommend",
        r"\bproblem solving\b",
        r"\bworkshop\b",
        r"\bclient\b",
        r"\bcommercial\b",
    ],
    "sales_bd": [
        r"\bclient\b",
        r"\brevenue\b",
        r"\bgrowth\b",
        r"\bnegotiat",
        r"\bpipeline\b",
        r"\bpartner",
        r"\baccount management\b",
    ],
    "operations": [
        r"\bprocess\b",
        r"\boperations\b",
        r"\bcoordinat",
        r"\befficiency\b",
        r"\bdocument",
        r"\bservice\b",
        r"\bstakeholder\b",
    ],
}


@dataclass
class CandidateFact:
    text: str
    source: str
    category: str
    score: float = 0.0


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def _compact(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip())


def _flatten_values(value: object) -> Iterable[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, (list, tuple, set)):
        items: List[str] = []
        for item in value:
            items.extend(_flatten_values(item))
        return items
    if isinstance(value, dict):
        items: List[str] = []
        for nested in value.values():
            items.extend(_flatten_values(nested))
        return items
    return [str(value)]


def _split_multi_fact_text(text: str) -> List[str]:
    compact = _compact(text)
    if not compact:
        return []
    if " | " in compact:
        parts = [part.strip() for part in compact.split("|")]
        return [part for part in parts if part]
    if "; " in compact and len(compact) > 140:
        parts = [part.strip() for part in compact.split(";")]
        return [part for part in parts if part]
    return [compact]


def _infer_category(text: str, source: str) -> str:
    low = text.lower()
    if any(re.search(pattern, low) for pattern in AVAILABILITY_PATTERNS):
        return "availability"
    if any(re.search(pattern, low) for pattern in QUALIFICATION_PATTERNS):
        return "qualification"
    if any(re.search(pattern, low) for pattern in IMPACT_PATTERNS):
        return "achievement"
    if source in {"experience", "projects", "external_fact"}:
        return "experience"
    return "profile"


def _base_profile_facts(profile: Dict[str, object]) -> List[CandidateFact]:
    facts: List[CandidateFact] = []
    identity_fields: List[Tuple[str, str]] = [
        ("name", "identity"),
        ("title", "identity"),
        ("location", "identity"),
    ]
    for key, category in identity_fields:
        value = _compact(profile.get(key, ""))
        if value:
            facts.append(CandidateFact(text=value, source=key, category=category))

    qualification_fields = [
        "education",
        "degree",
        "major",
        "graduation_date",
        "certifications",
        "qualifications",
    ]
    for key in qualification_fields:
        for value in _flatten_values(profile.get(key)):
            for fact in _split_multi_fact_text(value):
                facts.append(CandidateFact(text=fact, source=key, category="qualification"))

    optional_fields = [
        ("work_authorization", "availability"),
        ("availability_note", "availability"),
        ("skills", "profile"),
        ("summary", "profile"),
        ("headline", "profile"),
        ("experience", "experience"),
        ("projects", "projects"),
        ("achievements", "achievement"),
    ]
    for key, fallback_category in optional_fields:
        for value in _flatten_values(profile.get(key)):
            for fact in _split_multi_fact_text(value):
                facts.append(
                    CandidateFact(
                        text=fact,
                        source=key,
                        category=_infer_category(fact, key) or fallback_category,
                    )
                )
    for item in profile.get("finance_capabilities", []) or []:
        if not isinstance(item, dict):
            continue
        wording = _compact(item.get("wording", ""))
        if wording:
            evidence = str(item.get("evidence", "")).strip().lower()
            score_boost = 0.8 if evidence == "strong" else 0.45 if evidence == "moderate" else 0.0
            facts.append(
                CandidateFact(
                    text=wording,
                    source="finance_capabilities",
                    category="experience",
                    score=score_boost,
                )
            )
    for item in profile.get("accounting_software_exposure", []) or []:
        if not isinstance(item, dict):
            continue
        wording = _compact(item.get("wording", ""))
        if wording:
            evidence = str(item.get("evidence", "")).strip().lower()
            score_boost = 0.65 if evidence == "strong" else 0.35 if evidence == "moderate" else 0.0
            facts.append(
                CandidateFact(
                    text=wording,
                    source="accounting_software_exposure",
                    category="profile",
                    score=score_boost,
                )
            )
    return facts


def _external_verified_facts(verified_facts: Sequence[str]) -> List[CandidateFact]:
    facts: List[CandidateFact] = []
    for raw in verified_facts:
        for fact in _split_multi_fact_text(raw):
            facts.append(
                CandidateFact(
                    text=fact,
                    source="external_fact",
                    category=_infer_category(fact, "external_fact"),
                )
            )
    return facts


def _should_include_availability(jd_signals: JDSignals) -> bool:
    jd_text = " ".join(
        [
            jd_signals.cleaned_text,
            jd_signals.target_title,
            " ".join(jd_signals.responsibilities),
        ]
    ).lower()
    return bool(
        re.search(
            r"\b(visa|work rights|work authorization|available|start immediately|immediate start|full[- ]time|part[- ]time|graduate)\b",
            jd_text,
        )
    )


def _relevance_terms(jd_signals: JDSignals) -> List[str]:
    raw_terms = (
        [jd_signals.role_family, jd_signals.target_title]
        + jd_signals.keywords
        + jd_signals.core_skills
        + jd_signals.tools
        + jd_signals.responsibilities[:8]
    )
    terms: List[str] = []
    for term in raw_terms:
        normalized = _norm(term)
        if not normalized or len(normalized) < 3:
            continue
        terms.append(normalized)
    return list(dict.fromkeys(terms))


def _dedupe_key(text: str) -> str:
    normalized = _norm(text)
    normalized = re.sub(r"\b(douglas mo|brisbane australia|australia)\b", "", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def _is_generic_branding(text: str) -> bool:
    low = text.lower()
    return any(re.search(pattern, low) for pattern in GENERIC_BRANDING_PATTERNS)


def _contains_any(text: str, patterns: Sequence[str]) -> bool:
    low = text.lower()
    return any(re.search(pattern, low) for pattern in patterns)


def _jd_match_score(text: str, relevance_terms: Sequence[str]) -> float:
    normalized = _norm(text)
    score = 0.0
    for term in relevance_terms:
        if term and term in normalized:
            score += 1.0 if len(term.split()) > 1 else 0.45
    return min(score, 6.0)


def _role_specific_score(text: str, role_family: str) -> float:
    patterns = ROLE_RELEVANCE_PATTERNS.get(role_family, [])
    boosts = ROLE_PRIORITIES.get(role_family, {}).get("boost", [])
    downranks = ROLE_PRIORITIES.get(role_family, {}).get("downrank", [])

    normalized = _norm(text)
    score = 0.0
    if any(re.search(pattern, normalized) for pattern in patterns):
        score += 2.2
    score += sum(0.55 for term in boosts if term in normalized)
    score -= sum(1.4 for term in downranks if term in normalized)
    return score


def _role_specific_penalty(text: str, source: str, category: str, role_family: str) -> float:
    normalized = _norm(text)
    penalty = 0.0
    if is_accounting_family(role_family) or role_family == "accounting":
        if any(term in normalized for term in ["ai", "machine learning", "llm", "agent", "computer vision", "deep learning"]):
            penalty += 3.8
        if source == "projects" and not any(
            term in normalized
            for term in ["finance", "account", "tax", "report", "forecast", "compliance", "excel"]
        ):
            penalty += 3.0
        if "bridge" in normalized or "unique blend" in normalized:
            penalty += 2.4
    if role_family == "data_analyst" and any(term in normalized for term in ["bookkeeping", "accounts payable", "accounts receivable"]):
        penalty += 1.8
    if "salary" in normalized or "expected salary" in normalized:
        penalty += 3.0
    if category == "profile" and len(normalized.split()) > 24:
        penalty += 1.5
    return penalty


def _qualification_score(text: str) -> float:
    return 2.0 if _contains_any(text, QUALIFICATION_PATTERNS) else 0.0


def _impact_score(text: str) -> float:
    low = text.lower()
    score = 0.0
    if any(re.search(pattern, low) for pattern in IMPACT_PATTERNS):
        score += 2.8
    if re.search(r"\b(revenue|profit|margin|cost|accuracy|compliance|forecast|tax|client|process)\b", low):
        score += 0.9
    return score


def _availability_score(text: str, include_availability: bool) -> float:
    if not _contains_any(text, AVAILABILITY_PATTERNS):
        return 0.0
    return 1.4 if include_availability else -1.5


def _source_score(source: str, category: str) -> float:
    if category == "achievement":
        return 1.6
    if category == "qualification":
        return 1.3
    if source in {"experience", "external_fact"}:
        return 1.1
    if category == "availability":
        return 0.2
    return 0.4


def _noise_penalty(text: str, category: str) -> float:
    low = text.lower()
    penalty = 0.0
    if _is_generic_branding(low):
        penalty += 3.0
    if "salary" in low or "expected salary" in low:
        penalty += 4.0
    if category == "profile" and len(low.split()) > 28:
        penalty += 1.8
    if low.count(",") >= 5 and category == "profile":
        penalty += 2.0
    return penalty


def _cap_by_category(facts: List[CandidateFact], budget: int, include_availability: bool) -> List[CandidateFact]:
    capped: List[CandidateFact] = []
    category_limits = {
        "identity": 3,
        "qualification": 2,
        "availability": 1 if include_availability else 0,
        "achievement": max(3, budget // 3),
        "experience": max(4, budget // 2),
        "profile": 1,
    }
    used_by_category: Dict[str, int] = {}
    for fact in facts:
        limit = category_limits.get(fact.category, 1)
        if used_by_category.get(fact.category, 0) >= limit:
            continue
        capped.append(fact)
        used_by_category[fact.category] = used_by_category.get(fact.category, 0) + 1
        if len(capped) >= budget:
            break
    return capped


def rank_profile_facts(profile: Dict[str, object], verified_facts: Sequence[str], jd_signals: JDSignals) -> List[CandidateFact]:
    """Rank candidate facts for prompt injection.

    Role-aware examples built into the ranking rules:
    - `accounting`: boost tax, reconciliation, reporting, audit, Excel; down-rank AI project facts.
    - `data_analyst`: boost SQL, Python, dashboards, automation; down-rank bookkeeping-only facts.
    - `sales_bd`: boost client, revenue, pipeline, negotiation; down-rank salary and weak branding text.
    """
    include_availability = _should_include_availability(jd_signals)
    relevance_terms = _relevance_terms(jd_signals)
    raw_facts = _base_profile_facts(profile) + _external_verified_facts(verified_facts)

    deduped: Dict[str, CandidateFact] = {}
    for fact in raw_facts:
        fact.text = _compact(fact.text)
        if not fact.text:
            continue
        dedupe_key = _dedupe_key(fact.text)
        if not dedupe_key or len(dedupe_key) < 3:
            continue

        fact.score = (
            fact.score
            + _jd_match_score(fact.text, relevance_terms)
            + _role_specific_score(fact.text, jd_signals.role_family)
            + _qualification_score(fact.text)
            + _impact_score(fact.text)
            + _availability_score(fact.text, include_availability)
            + _source_score(fact.source, fact.category)
            - _noise_penalty(fact.text, fact.category)
            - _role_specific_penalty(fact.text, fact.source, fact.category, jd_signals.role_family)
        )

        if fact.category == "identity":
            fact.score += 4.5
        if fact.category == "qualification" and "graduate" in jd_signals.target_title.lower():
            fact.score += 1.4
        if len(fact.text.split()) <= 2 and fact.category not in {"identity", "qualification"}:
            fact.score -= 1.2

        existing = deduped.get(dedupe_key)
        if existing is None or fact.score > existing.score:
            deduped[dedupe_key] = fact

    ranked = sorted(
        deduped.values(),
        key=lambda fact: (
            fact.score,
            fact.category in {"achievement", "experience"},
            fact.source == "external_fact",
            len(fact.text),
        ),
        reverse=True,
    )
    return ranked


def filter_profile_facts(profile: Dict[str, object], verified_facts: Sequence[str], jd_signals: JDSignals) -> List[str]:
    """Return a compact, high-signal fact set for resume and cover-letter prompts."""
    budget = ROLE_FACT_BUDGETS.get(jd_signals.role_family, 10)
    include_availability = _should_include_availability(jd_signals)
    ranked = rank_profile_facts(profile, verified_facts, jd_signals)
    selected = _cap_by_category(ranked, budget=budget, include_availability=include_availability)
    return [fact.text for fact in selected if fact.score > 0.25][:budget]
