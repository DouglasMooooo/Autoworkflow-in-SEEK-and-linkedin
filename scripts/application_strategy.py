from __future__ import annotations

import re
from typing import Dict, List

from generation_models import ATSResult, JDSignals, ResumeAnalysis, StrategyDecision


TIER_1_FAMILIES = {"business_analyst", "data_analyst"}
TIER_2_FAMILIES = {"accounting_reporting", "accounting_systems"}

TIER_1_TITLE_HINTS = [
    "business analyst",
    "data analyst",
    "product analyst",
    "implementation consultant",
    "systems analyst",
]
TIER_2_TITLE_HINTS = [
    "assistant accountant",
    "costing analyst",
    "cost accountant",
    "financial analyst",
    "treasury analyst",
    "finance officer",
]
TIER_3_TITLE_HINTS = [
    "accounts assistant",
    "accounts payable",
    "accounts receivable",
    "finance support",
    "admin",
    "bookkeeper",
]


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def _contains_any(text: str, terms: List[str]) -> bool:
    low = _norm(text)
    return any(_norm(term) in low for term in terms if term)


def classify_target_tier(job_title: str, jd_signals: JDSignals) -> str:
    title = _norm(job_title)
    if _contains_any(title, TIER_1_TITLE_HINTS):
        return "Tier 1"
    if _contains_any(title, TIER_2_TITLE_HINTS):
        return "Tier 2"
    if _contains_any(title, TIER_3_TITLE_HINTS):
        return "Tier 3"
    if jd_signals.role_family in TIER_1_FAMILIES:
        return "Tier 1"
    if jd_signals.role_family in TIER_2_FAMILIES:
        return "Tier 2"
    return "Tier 3"


def suggest_positioning(jd_signals: JDSignals) -> str:
    role = jd_signals.role_family
    if role == "data_analyst":
        return "Data"
    if role in {"accounting_reporting", "accounting_systems"}:
        return "Finance"
    if role == "business_analyst":
        return "Hybrid"
    if any(tool in " ".join(jd_signals.tools).lower() for tool in ["power bi", "sql", "python"]) and any(
        kw in " ".join(jd_signals.keywords).lower() for kw in ["reporting", "finance", "analysis", "cost", "commercial"]
    ):
        return "Hybrid"
    return "Finance" if any("account" in _norm(k) or "finance" in _norm(k) for k in jd_signals.keywords) else "Hybrid"


def _scale(value: float, low: float = 0.0, high: float = 100.0) -> float:
    return max(low, min(high, round(value, 1)))


def _experience_alignment(analysis: ResumeAnalysis) -> float:
    keep_count = len(analysis.keep)
    boost_count = len(analysis.boost)
    downplay_count = len(analysis.downplay)
    raw = 45.0 + keep_count * 8.0 + boost_count * 4.5 - downplay_count * 2.0
    if analysis.missing_keywords:
        raw -= min(len(analysis.missing_keywords), 8) * 2.0
    return _scale(raw)


def _career_value(job_title: str, target_tier: str, jd_signals: JDSignals) -> float:
    score = {"Tier 1": 92.0, "Tier 2": 74.0, "Tier 3": 56.0}[target_tier]
    title = _norm(job_title)
    if _contains_any(title, ["graduate", "junior", "entry", "analyst", "consultant"]):
        score += 4.0
    if _contains_any(title, ["accounts payable", "accounts receivable", "accounts assistant", "admin"]):
        score -= 8.0
    if jd_signals.role_family == "data_analyst":
        score += 4.0
    if jd_signals.role_family in {"accounting_reporting", "accounting_systems"}:
        score -= 1.5
    return _scale(score)


def _interview_probability(skill_match: float, experience_alignment: float, ats_result: ATSResult, target_tier: str) -> float:
    base = (skill_match * 0.35) + (experience_alignment * 0.30) + (ats_result.score * 0.35)
    if target_tier == "Tier 1":
        base -= 4.0
    if target_tier == "Tier 3":
        base += 5.0
    if ats_result.missing_tier1_keywords:
        base -= min(len(ats_result.missing_tier1_keywords), 4) * 3.5
    return _scale(base)


def _priority_level(overall: float, target_tier: str) -> str:
    if overall >= 78:
        return "High"
    if overall >= 60:
        return "Medium" if target_tier != "Tier 1" else "High"
    return "Low"


def _apply_decision(overall: float, interview_probability: float, jd_signals: JDSignals) -> str:
    if overall < 38 or interview_probability < 32:
        return "Skip"
    if jd_signals.role_family == "sales_bd" and overall < 55:
        return "Skip"
    return "Apply"


def _cv_focus(positioning: str, jd_signals: JDSignals) -> str:
    if positioning == "Data":
        return "Data / BA"
    if positioning == "Finance":
        return "Finance"
    if jd_signals.role_family == "business_analyst":
        return "Hybrid"
    return "Hybrid"


def _cover_tone(positioning: str) -> str:
    if positioning == "Data":
        return "Data"
    if positioning == "Finance":
        return "Finance"
    return "Hybrid"


def _talking_points(job_title: str, jd_signals: JDSignals, selected_profile_facts: List[str]) -> List[str]:
    points: List[str] = []
    title = job_title.strip() or jd_signals.target_title
    points.append(f"Position yourself as a realistic fit for {title} by linking prior work to the role's day-to-day workflows.")
    if jd_signals.tools:
        points.append(f"Be ready to discuss practical use of {', '.join(jd_signals.tools[:3])} in business or finance reporting contexts.")
    if jd_signals.responsibilities:
        points.append(f"Prepare one concrete example that matches this responsibility: {jd_signals.responsibilities[0]}.")
    if selected_profile_facts:
        points.append(f"Lead with this verified evidence: {selected_profile_facts[0]}.")
    return points[:4]


def _recruiter_message(job_title: str, company: str, positioning: str) -> str:
    return (
        f"Hi, I have applied for the {job_title} role at {company}. "
        f"My background combines {positioning.lower()}-relevant experience across finance, reporting, and process improvement, "
        f"and I would welcome the chance to discuss how I could contribute quickly."
    )


def build_application_strategy(
    *,
    job_title: str,
    company: str,
    jd_signals: JDSignals,
    analysis: ResumeAnalysis,
    ats_result: ATSResult,
    selected_profile_facts: List[str],
) -> StrategyDecision:
    target_tier = classify_target_tier(job_title, jd_signals)
    positioning = suggest_positioning(jd_signals)

    skill_match_score = _scale(ats_result.score)
    experience_alignment_score = _experience_alignment(analysis)
    career_value_score = _career_value(job_title, target_tier, jd_signals)
    interview_probability_score = _interview_probability(skill_match_score, experience_alignment_score, ats_result, target_tier)

    overall_match_score = _scale(
        (skill_match_score * 0.42)
        + (experience_alignment_score * 0.23)
        + (career_value_score * 0.20)
        + (interview_probability_score * 0.15)
    )

    priority_level = _priority_level(overall_match_score, target_tier)
    apply_decision = _apply_decision(overall_match_score, interview_probability_score, jd_signals)
    cv_focus = _cv_focus(positioning, jd_signals)
    cover_letter_tone = _cover_tone(positioning)

    notes = [
        f"Target tier: {target_tier}",
        f"Suggested positioning: {positioning}",
        f"ATS score: {ats_result.score}",
    ]
    if ats_result.missing_tier1_keywords:
        notes.append(f"Missing tier-1 signals: {', '.join(ats_result.missing_tier1_keywords[:3])}")

    return StrategyDecision(
        target_tier=target_tier,
        priority_level=priority_level,
        apply_decision=apply_decision,
        suggested_positioning=positioning,
        cv_focus=cv_focus,
        cover_letter_tone=cover_letter_tone,
        skill_match_score=skill_match_score,
        experience_alignment_score=experience_alignment_score,
        career_value_score=career_value_score,
        interview_probability_score=interview_probability_score,
        overall_match_score=overall_match_score,
        recruiter_outreach_message=_recruiter_message(job_title, company, positioning),
        interview_talking_points=_talking_points(job_title, jd_signals, selected_profile_facts),
        decision_notes=notes,
    )
