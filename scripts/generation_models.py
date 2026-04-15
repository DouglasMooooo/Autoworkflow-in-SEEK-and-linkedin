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
class GenerationResult:
    role_family: str
    role_lock: str
    jd_signals: JDSignals
    analysis: ResumeAnalysis
    selected_profile_facts: List[str]
    resume_text: str
    cover_letter_text: str
    ats_result: ATSResult
    iteration_count: int
    skip_reason: str = ""
