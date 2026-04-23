from __future__ import annotations

from typing import Callable, Dict, List, Optional

from application_strategy import build_application_strategy
from ats_scorer import score_resume
from cover_letter_generator import generate_cover_letter
from finance_signal_injector import (
    finance_signal_prompt_facts,
    inject_finance_baseline_signals,
    resolve_accounting_subpersona,
    should_force_accounting_persona,
)
from generation_models import GenerationResult
from jd_parser import parse_jd
from profile_filter import filter_profile_facts
from resume_analyzer import analyze_resume, parse_base_resume_text
from resume_generator import generate_role_locked_resume, targeted_refinement_inputs


LLMCaller = Callable[[str], str]


def detect_hard_ineligibility(candidate_profile: Dict[str, object], jd_text: str) -> str:
    """Return a skip reason when the JD has hard work-rights constraints the candidate does not meet."""
    low = str(jd_text or "").lower()
    if not low:
        return ""

    requires_citizenship_or_pr = any(
        phrase in low
        for phrase in [
            "australian citizen",
            "australian or new zealand citizen",
            "permanent resident",
        ]
    )
    if not requires_citizenship_or_pr:
        return ""

    work_auth = str(candidate_profile.get("work_authorization", "")).lower()
    if any(term in work_auth for term in ["student visa", "24h/week", "24 hours", "485 graduate visa", "require employer sponsorship"]):
        return "Skipped due to hard work-rights requirement mismatch in JD."
    return ""


def optimize_application(
    *,
    job_title: str,
    company: str,
    location: str,
    jd_text: str,
    base_resume_text: str,
    candidate_profile: Dict[str, object],
    verified_facts: List[str],
    llm_caller: Optional[LLMCaller] = None,
    score_threshold: float = 85.0,
) -> GenerationResult:
    """Run the role-aware generation pipeline for a single job.

    The orchestrator owns decision order:
    JD parsing -> role classification -> resume analysis -> profile fact filtering
    -> resume generation -> ATS scoring -> one targeted refinement -> cover letter generation.
    """
    skip_reason = detect_hard_ineligibility(candidate_profile, jd_text)
    if skip_reason:
        jd_signals = parse_jd(job_title, jd_text)
        base_resume_data = parse_base_resume_text(base_resume_text)
        analysis = analyze_resume(base_resume_text, jd_signals)
        skip_ats = score_resume("", jd_signals, analysis)
        return GenerationResult(
            role_family=jd_signals.role_family,
            role_lock=job_title,
            jd_signals=jd_signals,
            analysis=analysis,
            selected_profile_facts=[],
            resume_text="",
            cover_letter_text="",
            ats_result=skip_ats,
            strategy=build_application_strategy(
                job_title=job_title,
                company=company,
                jd_signals=jd_signals,
                analysis=analysis,
                ats_result=skip_ats,
                selected_profile_facts=[],
            ),
            iteration_count=0,
            skip_reason=skip_reason,
        )

    jd_signals = parse_jd(job_title, jd_text)
    if should_force_accounting_persona(jd_signals):
        jd_signals.role_family = resolve_accounting_subpersona(jd_signals)
    base_resume_data = parse_base_resume_text(base_resume_text)
    analysis = analyze_resume(base_resume_text, jd_signals)
    selected_profile_facts = filter_profile_facts(candidate_profile, verified_facts, jd_signals)
    finance_signals = inject_finance_baseline_signals(selected_profile_facts, candidate_profile, verified_facts, jd_signals)
    finance_prompt_facts = finance_signal_prompt_facts(finance_signals)
    if finance_prompt_facts:
        selected_profile_facts = list(dict.fromkeys(finance_prompt_facts + selected_profile_facts))

    resume_text = generate_role_locked_resume(
        base_resume_text=base_resume_text,
        base_resume_data=base_resume_data,
        jd_signals=jd_signals,
        analysis=analysis,
        selected_profile_facts=selected_profile_facts,
        finance_signals=finance_signals,
        llm_caller=llm_caller,
    )
    ats_result = score_resume(resume_text, jd_signals, analysis)
    iteration_count = 1

    if ats_result.score < score_threshold:
        refinement_targets = targeted_refinement_inputs(ats_result, jd_signals)
        refined_resume = generate_role_locked_resume(
            base_resume_text=base_resume_text,
            base_resume_data=base_resume_data,
            jd_signals=jd_signals,
            analysis=analysis,
            selected_profile_facts=selected_profile_facts,
            finance_signals=finance_signals,
            llm_caller=llm_caller,
            refinement_targets=refinement_targets,
        )
        refined_score = score_resume(refined_resume, jd_signals, analysis)
        if refined_score.score >= ats_result.score:
            resume_text = refined_resume
            ats_result = refined_score
        iteration_count = 2

    cover_letter_text = generate_cover_letter(
        job_title=job_title,
        company=company,
        location=location,
        jd_signals=jd_signals,
        resume_text=resume_text,
        selected_profile_facts=selected_profile_facts,
        llm_caller=llm_caller,
    )
    strategy = build_application_strategy(
        job_title=job_title,
        company=company,
        jd_signals=jd_signals,
        analysis=analysis,
        ats_result=ats_result,
        selected_profile_facts=selected_profile_facts,
    )

    return GenerationResult(
        role_family=jd_signals.role_family,
        role_lock=job_title,
        jd_signals=jd_signals,
        analysis=analysis,
        selected_profile_facts=selected_profile_facts,
        resume_text=resume_text,
        cover_letter_text=cover_letter_text,
        ats_result=ats_result,
        strategy=strategy,
        iteration_count=iteration_count,
        skip_reason="",
    )
