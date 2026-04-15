from __future__ import annotations

import re
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from finance_signal_injector import FinanceSignalEvidence
from generation_models import ATSResult, JDSignals, ResumeAnalysis, ResumeRole
from role_classifier import is_accounting_family


LLMCaller = Callable[[str], str]

BANNED_SUMMARY_PHRASES = [
    "professional aligned to",
    "candidate aligned to",
    "results-driven individual aligned to",
    "practical experience in key responsibilities",
]

accounting_title_priority = True

ROLE_SKILL_MAP = {
    "accounting": [
        ("Accounting & Reporting", ["Month-End Reporting", "Reconciliations", "Tax Compliance", "Financial Controls"]),
        ("Systems & Analysis", ["Excel (Reporting, Reconciliations, Forecasting)", "ERP / SAP Systems", "Accounting Systems"]),
        ("Delivery Strengths", ["Attention to Detail", "Client Support", "Compliance Documentation"]),
    ],
    "accounting_reporting": [
        ("Accounting & Reporting", ["Financial Reporting", "Reconciliations", "Tax Compliance", "Month-End Support", "Excel Reporting"]),
        ("Controls & Analysis", ["Excel (Reporting, Reconciliations, Forecasting)", "Financial Statements", "Forecasting and Reporting Analysis"]),
        ("Working Style", ["Attention to Detail", "Compliance Documentation", "Accuracy and Deadline Discipline"]),
    ],
    "accounting_systems": [
        ("Finance Operations", ["Accounts Payable", "Accounts Receivable", "Invoicing", "Reconciliations", "Month-End Support"]),
        ("Systems & Reporting", ["Xero / MYOB / QuickBooks Exposure", "ERP / SAP Systems", "P/L Reporting", "Accounting Systems"]),
        ("Working Style", ["Accuracy", "Process Discipline", "Finance Operations Support"]),
    ],
    "data_analyst": [
        ("Analytics & Reporting", ["Data Analysis", "Dashboard Development", "Reporting", "Trend Analysis", "Data Quality"]),
        ("Tools", ["Power BI (Dashboard Development, KPI Tracking)", "Excel (Reporting, Pivot Tables, VLOOKUP)", "SQL (Data Extraction and Analysis)", "Python (Data Preparation and Automation)"]),
        ("Stakeholder Delivery", ["Stakeholder Reporting", "Insight Delivery", "Process and Reporting Support"]),
    ],
    "business_analyst": [
        ("Business Analysis", ["Process Improvement", "Workflow Documentation", "Requirements Gathering", "Operational Reporting"]),
        ("Tools & Systems", ["Excel (Reporting and Analysis)", "Power BI (Dashboard and KPI Reporting)", "CRM / ERP Systems", "Process Mapping and Documentation"]),
        ("Stakeholder Support", ["Stakeholder Communication", "Cross-Functional Coordination", "Problem Solving", "Documentation"]),
    ],
    "business_graduate": [
        ("Commercial Support", ["Reporting", "Coordination", "Customer Support", "Commercial Exposure"]),
        ("Execution Tools", ["Excel (Reporting and Administration)", "Power BI (Basic Dashboard Support)", "CRM / Administrative Systems"]),
        ("Working Style", ["Adaptability", "Cross-Functional Support", "Communication", "Learning Agility"]),
    ],
    "sales_bd": [
        ("Client & Commercial Delivery", ["Client Engagement", "Relationship Building", "Revenue Support", "Customer Communication"]),
        ("Tools", ["Excel (Reporting and Analysis)", "CRM Systems", "Pipeline and Activity Tracking"]),
        ("Working Style", ["Communication", "Commercial Awareness", "Execution", "Stakeholder Coordination"]),
    ],
}

weak_verbs = ["supported", "assisted", "worked on", "helped"]

strong_map = {
    "supported": "delivered",
    "assisted": "executed",
    "worked on": "led",
    "helped": "drove",
}


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def _human_join(items: Sequence[str]) -> str:
    values = [str(item).strip() for item in items if str(item).strip()]
    if not values:
        return ""
    if len(values) == 1:
        return values[0]
    if len(values) == 2:
        return f"{values[0]} and {values[1]}"
    return f"{', '.join(values[:-1])}, and {values[-1]}"


def _is_reporting_accounting(role_family: str) -> bool:
    return role_family in {"accounting", "accounting_reporting"}


def _is_systems_accounting(role_family: str) -> bool:
    return role_family == "accounting_systems"


def _pick_achievement_facts(selected_profile_facts: List[str], role_family: str) -> List[str]:
    scored: List[Tuple[float, str]] = []
    for fact in selected_profile_facts:
        text = str(fact).strip()
        low = text.lower()
        if len(text) > 180:
            continue
        if low.startswith(("i'm ", "i am ", "analytically-driven professional", "business-focused technologist", "cross-functional expertise")):
            continue
        score = 0.0
        if re.search(r"(\$|\d+%|\d+\+|accuracy|roi|revenue|forecast|churn|dashboard|tax|compliance|reporting)", low):
            score += 3.5
        if role_family == "data_analyst" and re.search(r"\b(power bi|sql|dashboard|analysis|data|report)\b", low):
            score += 2.6
        if role_family == "business_analyst" and re.search(r"\b(process|workflow|stakeholder|report|operations|commercial|support)\b", low):
            score += 2.4
        if role_family == "business_analyst" and re.search(r"\b(revenue|dashboard|roi|forecast|churn|reporting|operations)\b", low):
            score += 2.2
        if role_family == "business_graduate" and re.search(r"\b(customer|commercial|operations|support|team|report)\b", low):
            score += 2.0
        if _is_reporting_accounting(role_family):
            if re.search(r"\b(financial reporting|financial statements|reporting)\b", low):
                score += 4.8
            if re.search(r"\b(reconciliation|reconciliations|bank reconciliation)\b", low):
                score += 4.4
            if re.search(r"\b(month[- ]end|close process|close-process|closing processes)\b", low):
                score += 4.0
            if re.search(r"\b(tax|compliance|vat|corporate tax)\b", low):
                score += 3.7
            if re.search(r"\b(forecast accuracy|forecasting|cash-flow forecast|forecast)\b", low):
                score += 3.3
            if re.search(r"\b(quickbooks|xero|myob|sap|erp|accounting systems)\b", low):
                score += 2.4
            if re.search(r"\b(tax|compliance|forecast|financial|account|reconciliation)\b", low):
                score += 2.5
        if _is_systems_accounting(role_family):
            if re.search(r"\b(accounts payable|ap)\b", low):
                score += 5.0
            if re.search(r"\b(accounts receivable|ar|aging|collections)\b", low):
                score += 4.8
            if re.search(r"\b(invoic|billing)\b", low):
                score += 4.6
            if re.search(r"\b(reconciliation|reconciliations|bank reconciliation)\b", low):
                score += 4.2
            if re.search(r"\b(month[- ]end|close process)\b", low):
                score += 4.0
            if re.search(r"\b(p&l|profit and loss|financial reporting)\b", low):
                score += 3.8
            if re.search(r"\b(quickbooks|xero|myob|sap|erp|accounting systems)\b", low):
                score += 4.5
        if role_family in {"business_analyst", "business_graduate"} and re.search(r"\b(ml|machine learning|rag|llm|digital twin|vector)\b", low):
            score -= 1.8
        if is_accounting_family(role_family) and re.search(r"\b(ai|ml|machine learning|rag|llm|digital twin|vector|mcp|prompt)\b", low):
            score -= 4.5
        if low.startswith(("certification:", "portfolio:", "linkedin:", "location preferences:", "travel availability:")):
            score -= 3.0
        scored.append((score, text))
    scored.sort(key=lambda item: item[0], reverse=True)
    return [text for score, text in scored if score > 1.0][:3]


def _accounting_summary_achievement_sentence(selected_profile_facts: List[str], role_family: str) -> str:
    if _is_systems_accounting(role_family):
        ordered_patterns = [
            r"\baccounts payable\b|\bap\b",
            r"\baccounts receivable\b|\bar\b|\baging\b|\bcollections\b",
            r"\binvoic|\bbilling\b",
            r"\breconcil",
            r"\bmonth[- ]end\b|\bclose process\b",
            r"\bp&l\b|\bprofit and loss\b|\bfinancial reporting\b",
            r"\bquickbooks\b|\bxero\b|\bmyob\b|\bsap\b|\berp\b",
        ]
        fallback = "My finance operations experience includes supporting AP/AR workflows, reconciliations, month-end processes, and accounting systems with strong accuracy discipline."
    else:
        ordered_patterns = [
        r"\bfinancial reporting\b|\bfinancial statements?\b|\bat-risk revenue\b",
        r"\breconcil",
        r"\bmonth[- ]end\b|\bclose process\b|\bclosing processes\b",
        r"\btax\b|\bvat\b|\bcompliance\b",
        r"\bforecast accuracy\b|\bcash-flow forecast\b|\bforecast\b",
        r"\bquickbooks\b|\bxero\b|\bmyob\b|\bsap\b|\berp\b",
        ]
        fallback = "My finance experience includes improving reporting accuracy, supporting reconciliations, and maintaining strong compliance discipline."
    ranked: List[Tuple[float, str]] = []
    for fact in selected_profile_facts:
        text = str(fact).strip()
        if not text or len(text) > 180:
            continue
        low = text.lower()
        score = 0.0
        for idx, pattern in enumerate(ordered_patterns):
            if re.search(pattern, low):
                score += 8 - idx
        if re.search(r"(\$|\d+%|\d+\+|accuracy|revenue|forecast|tax filings|reporting cycle)", low):
            score += 3.5
        if re.search(r"\b(ai|ml|machine learning|rag|llm|digital twin|vector|mcp|prompt)\b", low):
            score -= 5.0
        ranked.append((score, text))

    ranked.sort(key=lambda item: item[0], reverse=True)
    top = [text for score, text in ranked if score > 2.5][:2]
    if top:
        if _is_systems_accounting(role_family):
            return f"My finance operations experience includes {_human_join(top)}."
        return f"My finance experience includes {_human_join(top)}."
    return fallback


def _persona_resume_title(base_resume_data: Dict[str, object], jd_signals: JDSignals, finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None) -> str:
    base_title = str(base_resume_data.get("title", "")).strip()
    if not is_accounting_family(jd_signals.role_family) or not accounting_title_priority:
        return base_title

    finance_terms = [item.summary_fragment for item in finance_signals or [] if item.summary_fragment][:3]
    if _is_systems_accounting(jd_signals.role_family):
        if "internal accountant" in jd_signals.target_title.lower():
            return "Internal Accountant | Reporting, Invoicing & Finance Systems"
        if "graduate" in jd_signals.target_title.lower():
            return "Graduate Accountant | AP, AR & Finance Systems"
        return "Junior Accountant | Reconciliations, Month-End & Accounting Systems"
    if "graduate" in jd_signals.target_title.lower():
        return "Graduate Accountant | Financial Reporting & Reconciliations"
    if any(term in finance_terms for term in ["month-end support", "reconciliations"]):
        return "Accountant | Reconciliations, Month-End & Reporting"
    return "Early-Career Accountant | Reporting, Compliance & Finance Operations"


def _summary_role_identity(role_family: str) -> str:
    if role_family == "data_analyst":
        return "Data analyst with a business analytics background and hands-on reporting experience"
    if role_family == "business_analyst":
        return "Business analyst with hands-on reporting, process support, and operational analysis experience"
    if role_family == "business_graduate":
        return "Business graduate candidate with reporting, coordination, and commercial support exposure"
    if _is_reporting_accounting(role_family):
        return "Early-career accountant with practical reporting, reconciliation, and compliance experience"
    if _is_systems_accounting(role_family):
        return "Early-career accountant with hands-on finance operations and accounting systems exposure"
    if role_family == "sales_bd":
        return "Commercially minded candidate with client-facing communication and customer support experience"
    return "Analytical professional with cross-functional business support experience"


def _finance_summary_identity(finance_signals: Sequence[FinanceSignalEvidence]) -> str:
    fragments = [item.summary_fragment for item in finance_signals if item.summary_fragment][:3]
    if not fragments:
        return "Early-career accounting professional with hands-on finance operations exposure"
    return f"Early-career accounting professional with hands-on experience across {_human_join(fragments)}"


def _summary_role_direction(jd_signals: JDSignals) -> str:
    if jd_signals.role_family == "data_analyst":
        return "I am targeting data analyst roles where I can use dashboards, reporting, and stakeholder-facing insight delivery to improve decision-making."
    if jd_signals.role_family == "business_analyst":
        return "I am targeting business analyst roles that need clear reporting, workflow support, stakeholder coordination, and practical process improvement."
    if jd_signals.role_family == "business_graduate":
        return "I am targeting graduate roles that combine reporting, coordination, adaptability, and cross-functional execution in a commercial environment."
    if _is_reporting_accounting(jd_signals.role_family):
        return "I am targeting accounting roles that value reporting accuracy, compliance discipline, Excel capability, and dependable support across finance processes."
    if _is_systems_accounting(jd_signals.role_family):
        return "I am targeting accounting systems and finance operations roles that value AP/AR support, reconciliations, month-end discipline, and practical accounting systems capability."
    if jd_signals.role_family == "sales_bd":
        return "I am targeting sales, customer, and commercial roles that value relationship building, communication, revenue support, and strong execution."
    return f"I am targeting {jd_signals.target_title} opportunities that value structured analysis, communication, and reliable execution."


def build_summary_lines(
    jd_signals: JDSignals,
    selected_profile_facts: List[str],
    refinement_targets: Optional[List[str]] = None,
    finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None,
) -> List[str]:
    achievements = _pick_achievement_facts(selected_profile_facts, jd_signals.role_family)
    if is_accounting_family(jd_signals.role_family) and finance_signals:
        sentence_one = _finance_summary_identity(finance_signals) + "."
        sentence_two = _accounting_summary_achievement_sentence(selected_profile_facts, jd_signals.role_family)
    else:
        sentence_one = _summary_role_identity(jd_signals.role_family) + "."
        sentence_two = "My experience includes measurable outcomes such as " + _human_join(achievements[:2]) + "." if achievements else (
            "I bring a track record of using data, reporting, and structured execution to support measurable business outcomes."
        )
    sentence_three = _summary_role_direction(jd_signals)
    if is_accounting_family(jd_signals.role_family) and finance_signals:
        software = [item.skill_label for item in finance_signals if item.signal in {"quickbooks", "sap", "xero", "myob"}][:2]
        if _is_systems_accounting(jd_signals.role_family):
            sentence_three = (
                "I am targeting finance operations and systems-focused accounting roles that value AP/AR support, invoicing, reconciliations, month-end discipline, "
                f"and practical exposure to {_human_join(software) if software else 'accounting systems and ERP workflows'}."
            )
        elif software:
            sentence_three = (
                "I am targeting accounting and finance operations roles that value reconciliations, reporting accuracy, compliance discipline, "
                f"and practical exposure to {_human_join(software)}."
            )
    if refinement_targets:
        clean_targets = [target for target in refinement_targets if target not in {"summary", "skills", "experience"}][:2]
        if clean_targets:
            sentence_three = sentence_three[:-1] + f", with particular strength in {_human_join(clean_targets)}."
    return [sentence_one, sentence_two, sentence_three]


def _upgrade_skill_label(skill: str, role_family: str) -> str:
    low = skill.lower()
    replacements = {
        "excel": "Excel (Reporting, Pivot Tables, VLOOKUP)",
        "power bi": "Power BI (Dashboard Development, KPI Tracking)",
        "sql": "SQL (Data Extraction and Analysis)",
        "python": "Python (Data Preparation and Automation)",
        "erp": "ERP / SAP Systems",
        "sap": "SAP / ERP Systems",
        "reporting": "Reporting and Performance Tracking",
        "workflow documentation": "Workflow Documentation and Process Mapping",
        "requirements gathering": "Requirements Gathering and Business Analysis",
        "process improvement": "Process Improvement and Operational Support",
        "tax compliance": "Tax Compliance and Regulatory Reporting",
        "reconciliations": "Reconciliations and Financial Accuracy",
        "crm": "CRM Systems",
        "stakeholder communication": "Stakeholder Engagement and Communication",
        "customer support": "Customer Support and Relationship Building",
    }
    return replacements.get(low, skill[:1].upper() + skill[1:])


def build_resume_prompt(
    base_resume_text: str,
    jd_signals: JDSignals,
    analysis: ResumeAnalysis,
    selected_profile_facts: List[str],
    finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None,
    refinement_targets: Optional[List[str]] = None,
) -> str:
    refinement_line = ""
    if refinement_targets:
        refinement_line = f"Targeted refinement focus: {', '.join(refinement_targets[:8])}."
    return f"""
You are a senior resume strategist and ATS optimization expert.

Task:
Improve the existing resume for ONE role only: {jd_signals.target_title}

Decision context:
- Role family: {jd_signals.role_family}
- Target profile:
  - {analysis.target_profile[0]}
  - {analysis.target_profile[1]}
  - {analysis.target_profile[2]}
- KEEP roles:
  - {'; '.join(role.header for role in analysis.keep) or 'None'}
- BOOST roles:
  - {'; '.join(role.header for role in analysis.boost) or 'None'}
- DOWNPLAY roles:
  - {'; '.join(role.header for role in analysis.downplay) or 'None'}
- High-value JD keywords:
  - {', '.join(jd_signals.keywords[:15])}
- Core JD skills:
  - {', '.join(jd_signals.core_skills[:10])}
- Tools:
  - {', '.join(jd_signals.tools[:8])}
- Qualifications:
  - {', '.join(jd_signals.qualifications[:8]) or 'None'}
- Domain terms:
  - {', '.join(jd_signals.domain_terms[:6]) or 'None'}
- Missing cleaned signals from current resume:
  - {', '.join(analysis.missing_keywords[:10]) or 'None'}
- Verified facts:
  - {' | '.join(selected_profile_facts[:18])}
- Finance baseline signals:
  - {', '.join(item.signal for item in (finance_signals or [])[:10]) or 'None'}
- Preserve ALL role headers exactly:
  - {' | '.join(analysis.preserved_headers)}

Strict rules:
- Never remove a role, company, date, or quantified achievement.
- Never merge multiple roles into generic blocks.
- Never add fake experience.
- Keep the resume ATS-safe and copy-ready.
- Use only sections: SUMMARY, SKILLS, EXPERIENCE, EDUCATION, AVAILABILITY.
- Do not include internal ATS notes or comments in the final resume.
- Never use banned summary phrasing such as "Professional aligned to", "Candidate aligned to", or "Results-driven individual aligned to".
- Summary must be 3 sentences:
  1. natural professional introduction
  2. one or two strongest relevant measurable achievements
  3. clear role direction and strengths relevant to the JD
- Downplay irrelevant content by wording and ordering, not by deleting roles.
- For accounting roles, surface finance operations language early and naturally: reconciliations, month-end support, financial reporting, tax compliance, accounting systems, and AP / AR / invoicing where supported.
- For accounting roles, the summary should sound like an accounting candidate, not a generic graduate.
- Business Analyst roles should prioritize process improvement, stakeholder support, workflow/documentation, reporting, and operational coordination.
- Data Analyst roles should prioritize dashboards, Power BI / Excel / SQL, reporting, trend analysis, and stakeholder-facing insights.
- Business Graduate roles should prioritize operational support, adaptability, reporting, coordination, commercial exposure, and growth potential.
- Sales / Customer / Commercial roles should prioritize client engagement, communication, relationship building, customer interaction, and revenue support.
- Rewrite bullets with ownership language. Avoid weak verbs like "supported", "assisted", "helped", and "worked on". Prefer direct ownership verbs such as "delivered", "executed", "drove", "managed", "developed", or "conducted" where accurate.
- For Business Analyst roles, translate implicit experience into business analyst language when truthful: stakeholder engagement, requirements gathering, process mapping, documentation, workflow improvement, and business solutions delivery.
- Keep business impact visible. If a bullet has a metric, connect it to efficiency, performance, decision-making, operational improvement, or commercial impact.
- Keep keywords natural and contextual.
- {refinement_line or 'First pass: optimize summary, skills, and role emphasis while preserving the base resume.'}

Return ONLY the final resume text.

JD:
{jd_signals.cleaned_text}

Base resume:
{base_resume_text}
""".strip()


def _skills_for_role(
    jd_signals: JDSignals,
    refinement_targets: Optional[List[str]] = None,
    finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None,
) -> List[Tuple[str, List[str]]]:
    categories = ROLE_SKILL_MAP.get(jd_signals.role_family) or ROLE_SKILL_MAP["business_graduate"]
    extra_terms = list(dict.fromkeys(jd_signals.core_skills + jd_signals.tools + jd_signals.qualifications))
    injected = [term for term in extra_terms if term][:5]
    output: List[Tuple[str, List[str]]] = []
    for idx, (heading, defaults) in enumerate(categories):
        merged = list(dict.fromkeys(defaults + [_upgrade_skill_label(item, jd_signals.role_family) for item in injected[idx: idx + 2]]))
        output.append((heading, merged[:6]))
    if is_accounting_family(jd_signals.role_family) and finance_signals and output:
        ordered_signals = sorted(finance_signals, key=lambda current: current.priority)
        if _is_systems_accounting(jd_signals.role_family):
            top_finance_labels = [
                item.skill_label
                for item in ordered_signals
                if item.signal in {"accounts_payable", "accounts_receivable", "invoicing", "reconciliations", "month_end_support", "p_and_l_reporting", "accounting_systems", "quickbooks", "xero", "myob", "sap"}
            ]
        else:
            top_finance_labels = [
                item.skill_label
                for item in ordered_signals
                if item.signal in {"financial_reporting", "reconciliations", "tax_compliance", "month_end_support", "p_and_l_reporting", "accounting_systems", "quickbooks", "xero", "myob", "sap"}
            ]
        output[0] = (output[0][0], list(dict.fromkeys(top_finance_labels[:4] + output[0][1]))[:6])
        if len(output) > 1:
            software_labels = [
                item.skill_label
                for item in ordered_signals
                if item.signal in {"quickbooks", "sap", "xero", "myob", "accounting_systems"}
            ]
            output[1] = (output[1][0], list(dict.fromkeys(software_labels[:4] + output[1][1]))[:6])
    if refinement_targets and output:
        output[0] = (
            output[0][0],
            list(
                dict.fromkeys(
                    output[0][1]
                    + [
                        _upgrade_skill_label(target, jd_signals.role_family)
                        for target in refinement_targets
                        if target not in {"summary", "skills", "experience"}
                    ]
                )
            )[:6],
        )
    return output


def _role_bullet_priority(role_family: str, bullet: str, finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None) -> float:
    low = bullet.lower()
    score = 0.0
    if re.search(r"(\$|\d+%|\d+\+|accuracy|roi|revenue|forecast|churn|saved|improved|reduced)", low):
        score += 2.2
    if role_family == "business_analyst" and re.search(r"\b(process|workflow|stakeholder|report|operations|coordination|documentation|support)\b", low):
        score += 3.2
    if role_family == "data_analyst" and re.search(r"\b(power bi|sql|dashboard|report|analysis|trend|data)\b", low):
        score += 3.2
    if role_family == "business_graduate" and re.search(r"\b(customer|support|team|operations|report|coordination|commercial)\b", low):
        score += 2.8
    if role_family == "sales_bd" and re.search(r"\b(client|customer|relationship|sales|revenue|commercial|communication)\b", low):
        score += 3.0
    if _is_reporting_accounting(role_family) and re.search(r"\b(tax|compliance|financial|reconciliation|ledger|forecast)\b", low):
        score += 3.0
    if _is_systems_accounting(role_family) and re.search(r"\b(accounts payable|accounts receivable|ap|ar|invoic|month[- ]end|p&l|profit and loss|quickbooks|xero|myob|sap|erp)\b", low):
        score += 3.6
    if is_accounting_family(role_family) and finance_signals:
        for signal in finance_signals:
            signal_text = signal.signal.replace("_", " ")
            if signal_text in low or signal.summary_fragment.lower() in low:
                score += 0.6
    if re.search(r"\b(ai|rag|llm|vector|digital twin|mcp)\b", low) and (is_accounting_family(role_family) or role_family == "business_graduate"):
        score -= 1.8
    return score


def _replace_weak_ownership_verbs(text: str) -> str:
    updated = text
    updated = re.sub(r"\bhelped with\b", "drove", updated, flags=re.I)
    updated = re.sub(r"\bassisted with\b", "executed", updated, flags=re.I)
    updated = re.sub(r"\bsupported with\b", "delivered", updated, flags=re.I)
    updated = re.sub(r"\bworked on\b", "led", updated, flags=re.I)
    for weak in weak_verbs:
        strong = strong_map[weak]
        updated = re.sub(rf"^\b{re.escape(weak)}\b", strong, updated, flags=re.I)
        updated = re.sub(rf"\b{re.escape(weak)}\b", strong, updated, flags=re.I)
    return updated


def _add_business_impact_phrase(text: str) -> str:
    low = text.lower()
    if any(term in low for term in ["decision-making", "performance tracking", "operational improvement", "commercial impact", "efficiency"]):
        return text
    if "dashboard" in low:
        return text.rstrip(".") + " to support decision-making and performance tracking."
    if "report" in low or "reporting" in low:
        return text.rstrip(".") + " to support business decision-making."
    if "process" in low or "workflow" in low:
        return text.rstrip(".") + " to improve efficiency and reduce operational friction."
    if "identified" in low and "issue" in low:
        return text.rstrip(".") + " and quantified the impact on business performance."
    return text


def _persona_rewrite_generic(text: str, role_family: str) -> str:
    updated = _replace_weak_ownership_verbs(text)
    low = updated.lower()

    if role_family == "business_analyst":
        if "reporting" in low and "stakeholder" not in low:
            updated = updated.rstrip(".") + " for stakeholder decision-making."
        if "process" in low and "documentation" not in low:
            updated = updated.rstrip(".") + " while improving process documentation and workflow clarity."
        if "analysis" in low and "requirements" not in low:
            updated = updated.rstrip(".") + " to inform requirements and business solutions delivery."
    elif role_family == "data_analyst":
        if "dashboard" in low and "power bi" not in low:
            updated = updated.rstrip(".") + " using Power BI and Excel."
        if "reporting" in low and "analysis" not in low:
            updated = updated.rstrip(".") + " through structured analysis and KPI tracking."
    elif _is_reporting_accounting(role_family):
        if "reporting" in low and "financial" not in low:
            updated = updated.rstrip(".") + " across financial reporting and compliance workflows."
        if "records" in low and "reconciliation" not in low:
            updated = updated.rstrip(".") + " with strong attention to reconciliation and reporting accuracy."
    elif _is_systems_accounting(role_family):
        if "reporting" in low and "financial" not in low and "p/l" not in low:
            updated = updated.rstrip(".") + " across finance operations reporting and reconciliation workflows."
        if "records" in low and "reconciliation" not in low:
            updated = updated.rstrip(".") + " with strong attention to reconciliation accuracy and systems discipline."
    elif role_family == "business_graduate":
        if "operations" in low and "execution" not in low:
            updated = updated.rstrip(".") + " while building execution discipline in a fast-moving environment."
        if "customer" in low and "communication" not in low:
            updated = updated.rstrip(".") + " through clear communication and day-to-day coordination."
    elif role_family == "sales_bd":
        if "customer" in low or "client" in low:
            updated = updated.rstrip(".") + " to strengthen relationships and support commercial outcomes."
        elif "reporting" in low:
            updated = updated.rstrip(".") + " while supporting client and revenue-facing decisions."

    return _add_business_impact_phrase(updated)


def _rewrite_bullet_contextually(bullet: str, role_family: str) -> str:
    text = str(bullet or "").strip()
    if not text:
        return text
    low = text.lower()

    if "built dashboards identifying" in low and "$1.8m" in low:
        if role_family == "data_analyst":
            return "Developed Power BI dashboards to identify $1.8M at-risk revenue and support decision-making."
        if role_family == "business_analyst":
            return "Analysed business data to identify $1.8M at-risk revenue, supporting stakeholder decision-making."
        if _is_reporting_accounting(role_family):
            return "Delivered financial reporting analysis that identified $1.8M in at-risk revenue and highlighted exposure for decision-making."
        if _is_systems_accounting(role_family):
            return "Delivered P/L and finance reporting analysis that identified $1.8M in at-risk revenue and highlighted exposure for action."
        if role_family == "business_graduate":
            return "Delivered reporting and analysis work that identified $1.8M in at-risk revenue for business decision-making."

    if "reduced churn from 22%" in low or "reduced distributor churn" in low:
        if role_family == "data_analyst":
            return "Analysed customer and revenue data through reporting dashboards, helping reduce churn from 22% to 16% and protect about $600K in revenue."
        if role_family == "business_analyst":
            return "Used business reporting and performance analysis to support actions that reduced churn from 22% to 16% and protected about $600K in revenue."
        if role_family == "business_graduate":
            return "Delivered commercial reporting that contributed to reducing churn from 22% to 16% and protecting about $600K in revenue."

    if "improved promotional roi from 8%" in low or "modeled promotions raising roi" in low:
        if role_family in {"data_analyst", "business_analyst"}:
            return "Applied Excel-based analysis and reporting to evaluate promotion performance, improving ROI from 8% to 23% and supporting an additional $85K in profit."
        if _is_reporting_accounting(role_family):
            return "Delivered commercial and financial reporting analysis that improved ROI from 8% to 23% and contributed to $85K in profit."
        if _is_systems_accounting(role_family):
            return "Delivered finance reporting analysis that improved ROI from 8% to 23% and strengthened profit visibility by $85K."

    if "developed forecasting model improving accuracy from 70% to 94%" in low or "improved cash-flow forecast accuracy" in low:
        if role_family == "data_analyst":
            return "Applied Excel-based forecasting and reporting to improve forecast accuracy from 70% to 94% and support better operational planning."
        if role_family == "business_analyst":
            return "Improved forecasting and reporting processes, increasing forecast accuracy from 70% to 94% to support business planning."
        if _is_reporting_accounting(role_family):
            return "Improved cash-flow forecasting and reporting accuracy from 70% to 94%, supporting stronger financial control and planning."
        if _is_systems_accounting(role_family):
            return "Improved cash-flow forecasting and reporting accuracy from 70% to 94%, supporting month-end control and finance operations planning."

    if "processed 1,000+ corporate tax filings" in low or "processed 1,000+ tax filings" in low:
        if _is_reporting_accounting(role_family):
            return "Processed 1,000+ tax filings with full compliance, supporting accurate reporting and month-end finance administration."
        if _is_systems_accounting(role_family):
            return "Processed 1,000+ tax filings with full compliance while maintaining accurate finance records and reporting workflows."
        if role_family == "business_analyst":
            return "Maintained accurate financial records and reporting processes across 1,000+ tax filings, supporting compliance and workflow discipline."
        if role_family == "business_graduate":
            return "Executed high-volume compliance and reporting work across 1,000+ tax filings, building strong execution and accuracy discipline."

    if "reduced reporting time by 35% via automation" in low:
        if role_family == "data_analyst":
            return "Used Excel automation to reduce reporting time by 35% and improve reporting efficiency."
        if role_family == "business_analyst":
            return "Improved workflow efficiency by automating reporting tasks and reducing reporting time by 35%."
        if _is_reporting_accounting(role_family):
            return "Applied Excel automation to reduce reporting cycle time by 35% and improve finance process efficiency."
        if _is_systems_accounting(role_family):
            return "Applied Excel automation to reduce reporting cycle time by 35% and improve finance operations efficiency."
        if role_family == "business_graduate":
            return "Delivered process improvement work that reduced reporting time by 35% through automation."

    if "managed financial records for 20+ clients" in low:
        if _is_reporting_accounting(role_family):
            return "Managed financial records for 20+ clients, supporting reconciliations, compliance, and reporting accuracy."
        if _is_systems_accounting(role_family):
            return "Managed financial records for 20+ clients, supporting reconciliations, invoicing discipline, and finance systems accuracy."
        if role_family == "business_analyst":
            return "Maintained accurate records for 20+ clients, supporting reporting, documentation, and stakeholder service."

    if "supported merchandising, replenishment, and checkout operations" in low:
        if role_family == "business_graduate":
            return "Delivered day-to-day store operations across merchandising, replenishment, and checkout, building coordination and execution capability in a high-volume environment."
        if role_family == "business_analyst":
            return "Managed day-to-day operations in a high-volume environment, building process discipline and cross-functional coordination capability."

    if "assisted stocktake and inventory verification" in low:
        if role_family in {"business_analyst", "data_analyst"}:
            return "Executed inventory verification and stocktake activities, improving data accuracy and reporting reliability."
        if role_family == "business_graduate":
            return "Executed stocktake and inventory verification, improving accuracy and operational support capability."

    if "delivered customer service while maintaining transaction accuracy" in low:
        if role_family == "business_graduate":
            return "Delivered customer support while maintaining transaction accuracy under time pressure."
        if role_family == "business_analyst":
            return "Managed customer-facing operations while maintaining process accuracy and service standards under time pressure."

    if "built ai-powered digital twin interview assistant" in low:
        if role_family == "data_analyst":
            return "Built a production analytics assistant with 95% accuracy and sub-2-second response time, strengthening data access and reporting support."
        if role_family == "business_analyst":
            return "Built a workflow support assistant with 95% accuracy and sub-2-second response time, improving information access and process support."
        if role_family == "business_graduate":
            return "Built a production support tool with 95% accuracy and sub-2-second response time, demonstrating adaptability and fast learning."

    if "designed rag architecture" in low:
        if role_family == "data_analyst":
            return "Designed structured data retrieval workflows to improve information access and reporting reliability."
        if role_family == "business_analyst":
            return "Designed structured workflow logic to improve information access, documentation, and process support."

    if "developed ai agents and mcp-based tooling for workflow automation" in low:
        if role_family in {"business_analyst", "business_graduate"}:
            return "Developed workflow automation tools that reduced manual effort and improved process consistency."
        if role_family == "data_analyst":
            return "Developed automation tools that improved reporting workflow efficiency and reduced manual effort."

    if "delivered 5+ ai solutions" in low:
        if role_family == "data_analyst":
            return "Delivered 5+ analytics and automation solutions, including dashboards and decision-support tools."
        if role_family == "business_analyst":
            return "Delivered 5+ analytics and workflow solutions to support reporting, documentation, and business decision-making."
        if role_family == "business_graduate":
            return "Delivered 5+ project outcomes across analytics and automation work, demonstrating learning agility and execution."

    if "optimised prompt and inference pipelines for performance" in low:
        if role_family in {"business_analyst", "business_graduate", "data_analyst"}:
            return "Optimised system performance and documentation to support reliable day-to-day use."
    return _persona_rewrite_generic(text, role_family)


def _sorted_roles(analysis: ResumeAnalysis, role_family: str, finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None) -> List[ResumeRole]:
    combined = analysis.keep + analysis.boost + analysis.downplay
    return sorted(
        combined,
        key=lambda role: (
            role.classification == "KEEP",
            role.relevance_score,
            max((_role_bullet_priority(role_family, bullet, finance_signals) for bullet in role.bullets), default=0.0),
        ),
        reverse=True,
    )


def _render_role(role: ResumeRole, role_family: str, finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None) -> List[str]:
    lines = [role.header]
    ranked_bullets = sorted(role.bullets, key=lambda bullet: _role_bullet_priority(role_family, bullet, finance_signals), reverse=True)
    for bullet in ranked_bullets:
        lines.append(f"- {_rewrite_bullet_contextually(bullet, role_family)}")
    return lines


def _sanitise_generated_resume(text: str, analysis: ResumeAnalysis) -> str:
    cleaned_lines: List[str] = []
    for raw in str(text or "").splitlines():
        line = raw.rstrip()
        low = line.lower()
        if any(phrase in low for phrase in BANNED_SUMMARY_PHRASES):
            continue
        cleaned_lines.append(line)
    cleaned_text = "\n".join(cleaned_lines).strip()
    for header in analysis.preserved_headers:
        if header not in cleaned_text:
            return ""
    return cleaned_text


def fallback_generate_resume(
    base_resume_data: Dict[str, object],
    jd_signals: JDSignals,
    analysis: ResumeAnalysis,
    selected_profile_facts: List[str],
    finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None,
    refinement_targets: Optional[List[str]] = None,
) -> str:
    lines: List[str] = []
    lines.append(str(base_resume_data.get("name", "")))
    lines.append(_persona_resume_title(base_resume_data, jd_signals, finance_signals))
    lines.append(str(base_resume_data.get("contact", "")))
    lines.append("LinkedIn | GitHub | Portfolio")
    lines.append("")
    lines.append("SUMMARY")
    lines.extend(build_summary_lines(jd_signals, selected_profile_facts, refinement_targets, finance_signals))
    lines.append("")
    lines.append("SKILLS")
    for heading, skills in _skills_for_role(jd_signals, refinement_targets, finance_signals):
        lines.append(heading)
        lines.append(f"- {_human_join(skills)}")
    lines.append("")
    lines.append("EXPERIENCE")
    for role in _sorted_roles(analysis, jd_signals.role_family, finance_signals):
        lines.extend(_render_role(role, jd_signals.role_family, finance_signals))
    if base_resume_data.get("education"):
        lines.append("")
        lines.append("EDUCATION")
        lines.extend([str(x) for x in base_resume_data.get("education", []) if str(x).strip()])
    lines.append("")
    lines.append("AVAILABILITY")
    availability = [str(x) for x in base_resume_data.get("availability", []) if str(x).strip()]
    if availability:
        lines.extend(availability)
    else:
        lines.append("Available based on confirmed candidate profile and visa timeline.")
    return "\n".join(lines).strip() + "\n"


def generate_role_locked_resume(
    base_resume_text: str,
    base_resume_data: Dict[str, object],
    jd_signals: JDSignals,
    analysis: ResumeAnalysis,
    selected_profile_facts: List[str],
    finance_signals: Optional[Sequence[FinanceSignalEvidence]] = None,
    llm_caller: Optional[LLMCaller] = None,
    refinement_targets: Optional[List[str]] = None,
) -> str:
    if llm_caller is not None:
        prompt = build_resume_prompt(
            base_resume_text=base_resume_text,
            jd_signals=jd_signals,
            analysis=analysis,
            selected_profile_facts=selected_profile_facts,
            finance_signals=finance_signals,
            refinement_targets=refinement_targets,
        )
        try:
            text = _sanitise_generated_resume(str(llm_caller(prompt)).strip(), analysis)
            if text:
                return text
        except Exception:
            pass
    return fallback_generate_resume(base_resume_data, jd_signals, analysis, selected_profile_facts, finance_signals, refinement_targets)


def targeted_refinement_inputs(ats_result: ATSResult, jd_signals: JDSignals) -> List[str]:
    targets: List[str] = []
    targets.extend(ats_result.missing_tier1_keywords[:4])
    for keyword in ats_result.missing_keywords:
        low = keyword.lower()
        if keyword in ats_result.missing_tier1_keywords:
            continue
        if any(noisy in low for noisy in ["about", "hew level", "superannuation", "benefits", "culture", "opportunity", "why join"]):
            continue
        if keyword in jd_signals.keywords or keyword in jd_signals.core_skills or keyword in jd_signals.tools:
            targets.append(keyword)
    targets.extend(ats_result.weak_sections)
    deduped: List[str] = []
    seen = set()
    for target in targets:
        key = _norm(target)
        if key and key not in seen:
            seen.add(key)
            deduped.append(target)
    return deduped[:8]
