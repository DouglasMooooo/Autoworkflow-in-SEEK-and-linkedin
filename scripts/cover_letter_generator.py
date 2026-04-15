from __future__ import annotations

from typing import Callable, List, Optional

from generation_models import JDSignals


LLMCaller = Callable[[str], str]


def build_cover_letter_prompt(job_title: str, company: str, location: str, jd_signals: JDSignals, resume_text: str, selected_profile_facts: List[str]) -> str:
    return f"""
You are a professional cover letter writer.

Write a concise, human, professional cover letter for:
- Role: {job_title}
- Company: {company}
- Location: {location}
- Role family: {jd_signals.role_family}

Rules:
- No internal ATS notes.
- No fake claims.
- Use only verified facts from the resume and selected profile facts.
- Keep it under 280 words.
- Sound natural, specific, and motivated.
- If the role is accounting_reporting, emphasize reporting, reconciliations, tax, compliance, attention to detail, and practical learning mindset.
- If the role is accounting_systems, emphasize AP/AR, invoicing, reconciliations, month-end support, accounting systems, and reliable finance operations execution.

Relevant JD signals:
- Keywords: {', '.join(jd_signals.keywords[:12])}
- Core skills: {', '.join(jd_signals.core_skills[:8])}

Selected profile facts:
- {' | '.join(selected_profile_facts[:12])}

Resume:
{resume_text}
""".strip()


def fallback_cover_letter(job_title: str, company: str, location: str, jd_signals: JDSignals, selected_profile_facts: List[str]) -> str:
    fact_line = selected_profile_facts[0] if selected_profile_facts else "I bring relevant experience and measurable results."
    return f"""{company}

Dear Hiring Manager,

I am applying for the {job_title} opportunity at {company}. My background in finance, reporting, and business analysis gives me a strong foundation for this role, and I am particularly interested in building practical experience in {jd_signals.role_family.replace('_', ' ')} work.

I bring experience that includes {fact_line}. Across my previous work, I have supported accurate reporting, process discipline, and measurable business outcomes while working carefully under deadlines and communicating clearly with stakeholders.

I would welcome the opportunity to contribute to your team in {location or 'this role'} and continue developing my capability in line with the requirements of the position.

Sincerely,
Douglas Mo
""".strip()


def generate_cover_letter(job_title: str, company: str, location: str, jd_signals: JDSignals, resume_text: str, selected_profile_facts: List[str], llm_caller: Optional[LLMCaller] = None) -> str:
    """Generate a separate cover letter without leaking ATS-internal notes."""
    if llm_caller is not None:
        prompt = build_cover_letter_prompt(job_title, company, location, jd_signals, resume_text, selected_profile_facts)
        try:
            text = str(llm_caller(prompt)).strip()
            if text:
                return text
        except Exception:
            pass
    return fallback_cover_letter(job_title, company, location, jd_signals, selected_profile_facts)
