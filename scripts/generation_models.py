from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class JDSignals:
    role_family: str
    target_title: str
    cleaned_text: str
    keywords: List[str]
    core_skills: List[str]
    responsibilities: List[str]
    tools: List[str]
    qualifications: List[str] = field(default_factory=list)
    domain_terms: List[str] = field(default_factory=list)


@dataclass
class ResumeRole:
    header: str
    bullets: List[str]
    classification: str
    relevance_score: float
    evidence: List[str] = field(default_factory=list)


@dataclass
class ResumeAnalysis:
    keep: List[ResumeRole]
    boost: List[ResumeRole]
    downplay: List[ResumeRole]
    missing_keywords: List[str]
    preserved_headers: List[str]
    target_profile: List[str]


@dataclass
class ATSResult:
    score: float
    matched_keywords: List[str]
    missing_keywords: List[str]
    matched_tier1_keywords: List[str]
    missing_tier1_keywords: List[str]
    weak_sections: List[str]
    improvements: List[str]
    debug: Dict[str, float] = field(default_factory=dict)


@dataclass
class StrategyDecision:
    target_tier: str
    priority_level: str
    apply_decision: str
    suggested_positioning: str
    cv_focus: str
    cover_letter_tone: str
    skill_match_score: float
    experience_alignment_score: float
    career_value_score: float
    interview_probability_score: float
    overall_match_score: float
    recruiter_outreach_message: str = ""
    interview_talking_points: List[str] = field(default_factory=list)
    decision_notes: List[str] = field(default_factory=list)


@dataclass
class GenerationResult:
    role_family: str
    role_lock: str
    jd_signals: JDSignals
    analysis: ResumeAnalysis
    selected_profile_facts: List[str]
    resume_text: str
    cover_letter_text: str
    ats_result: ATSResult
    strategy: StrategyDecision
    iteration_count: int
    skip_reason: str = ""
