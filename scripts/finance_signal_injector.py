from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Sequence

from generation_models import JDSignals
from role_classifier import is_accounting_family


FINANCE_TARGET_PATTERNS = [
    r"\baccountant\b",
    r"\baccounting\b",
    r"\bfinance operations\b",
    r"\bfinance officer\b",
    r"\baccounts officer\b",
    r"\baccounts assistant\b",
    r"\bbookkeep",
    r"\bgraduate accountant\b",
    r"\baccounts\b",
]


@dataclass
class FinanceSignalEvidence:
    signal: str
    level: str
    wording: str
    skill_label: str
    summary_fragment: str
    priority: int


FINANCE_SIGNAL_CONFIG: Dict[str, Dict[str, object]] = {
    "reconciliations": {
        "priority": 1,
        "patterns": [r"\breconcil", r"\bbank reconciliation", r"\bcash position", r"\b8 bank accounts?\b"],
        "strong_wording": "Experience with reconciliations across cash management, reporting, and finance control workflows.",
        "moderate_wording": "Exposure to reconciliation workflows and finance control processes.",
        "skill_label": "Reconciliations and Finance Controls",
        "summary_fragment": "reconciliations",
    },
    "month_end_support": {
        "priority": 2,
        "patterns": [r"\bmonth[- ]end", r"\byear[- ]end", r"\bclose process", r"\bclosing processes\b"],
        "strong_wording": "Experience with month-end support, close-process discipline, and reporting preparation.",
        "moderate_wording": "Exposure to month-end support and reporting preparation workflows.",
        "skill_label": "Month-End Support and Reporting Preparation",
        "summary_fragment": "month-end support",
    },
    "financial_reporting": {
        "priority": 3,
        "patterns": [r"\bfinancial reporting\b", r"\bfinancial statements?\b", r"\bp&l\b", r"\bprofit and loss\b", r"\breporting\b"],
        "strong_wording": "Delivered financial reporting, performance tracking, and finance analysis across operational workflows.",
        "moderate_wording": "Experience with financial reporting and finance performance tracking.",
        "skill_label": "Financial Reporting and P/L Tracking",
        "summary_fragment": "financial reporting",
    },
    "p_and_l_reporting": {
        "priority": 4,
        "patterns": [r"\bp&l\b", r"\bprofit and loss\b", r"\bcontribution margin\b", r"\bmargin analysis\b"],
        "strong_wording": "Experience with profit and loss reporting, performance tracking, and finance decision support.",
        "moderate_wording": "Exposure to P/L reporting and finance performance tracking.",
        "skill_label": "P/L Reporting and Performance Analysis",
        "summary_fragment": "P/L reporting",
    },
    "tax_compliance": {
        "priority": 5,
        "patterns": [r"\btax\b", r"\bvat\b", r"\bcorporate tax\b", r"\bwithholding\b", r"\bcompliance\b"],
        "strong_wording": "Delivered tax compliance work across high-volume filings with accuracy and reporting discipline.",
        "moderate_wording": "Experience with tax compliance and regulatory reporting support.",
        "skill_label": "Tax Compliance and Regulatory Reporting",
        "summary_fragment": "tax compliance",
    },
    "accounts_payable": {
        "priority": 6,
        "patterns": [r"\baccounts payable\b", r"\ba/p\b", r"\bpayment scheduling\b", r"\ba p schedules?\b"],
        "strong_wording": "Experience with accounts payable coordination, payment scheduling, and finance operations support.",
        "moderate_wording": "Exposure to accounts payable workflows and payment coordination.",
        "skill_label": "Accounts Payable (AP)",
        "summary_fragment": "accounts payable",
    },
    "accounts_receivable": {
        "priority": 7,
        "patterns": [r"\baccounts receivable\b", r"\ba/r\b", r"\bar aging\b", r"\ba r aging\b", r"\bcollections\b", r"\bpayment patterns\b"],
        "strong_wording": "Experience with accounts receivable tracking, aging visibility, and customer account workflows.",
        "moderate_wording": "Exposure to accounts receivable workflows and customer account tracking.",
        "skill_label": "Accounts Receivable (AR)",
        "summary_fragment": "accounts receivable",
    },
    "invoicing": {
        "priority": 8,
        "patterns": [r"\binvoic", r"\bbilling\b"],
        "strong_wording": "Experience with invoicing, billing coordination, and finance operations administration.",
        "moderate_wording": "Exposure to invoicing and billing workflows within finance operations.",
        "skill_label": "Invoicing and Billing Coordination",
        "summary_fragment": "invoicing",
    },
    "accounting_systems": {
        "priority": 9,
        "patterns": [r"\berp\b", r"\bsap\b", r"\bquickbooks\b", r"\bxero\b", r"\bmyob\b", r"\baccounting systems?\b"],
        "strong_wording": "Experience with accounting systems and ERP-supported finance workflows.",
        "moderate_wording": "Exposure to accounting systems and ERP-based finance processes.",
        "skill_label": "Accounting Systems and ERP Workflows",
        "summary_fragment": "accounting systems exposure",
    },
}


SOFTWARE_CONFIG = {
    "xero": {
        "skill_label": "Xero (SME Accounting Workflow Exposure)",
        "summary_fragment": "Xero exposure",
    },
    "myob": {
        "skill_label": "MYOB (SME Accounting Workflow Exposure)",
        "summary_fragment": "MYOB exposure",
    },
    "quickbooks": {
        "skill_label": "QuickBooks (Finance Operations and Reporting)",
        "summary_fragment": "QuickBooks exposure",
    },
    "sap": {
        "skill_label": "SAP / ERP Systems",
        "summary_fragment": "ERP systems exposure",
    },
}


def is_finance_target(jd_signals: JDSignals) -> bool:
    title_and_role = f"{jd_signals.target_title} {jd_signals.role_family}".lower()
    if is_accounting_family(jd_signals.role_family):
        return True
    return any(re.search(pattern, title_and_role) for pattern in FINANCE_TARGET_PATTERNS)


def should_force_accounting_persona(jd_signals: JDSignals) -> bool:
    title = str(jd_signals.target_title or "").lower()
    return any(
        re.search(pattern, title)
        for pattern in [
            r"\bgraduate accountant\b",
            r"\bjunior accountant\b",
            r"\baccountant\b",
            r"\bbookkeeper\b",
            r"\baccounts officer\b",
            r"\baccounts assistant\b",
            r"\bfinance officer\b",
            r"\bfinance operations\b",
        ]
    )


def resolve_accounting_subpersona(jd_signals: JDSignals) -> str:
    text = f"{jd_signals.target_title}\n{jd_signals.cleaned_text}".lower()
    systems_hits = sum(
        1
        for pattern in [
            r"\baccounts payable\b",
            r"\baccounts receivable\b",
            r"\binvoic",
            r"\bbilling\b",
            r"\bp&l\b",
            r"\bprofit and loss\b",
            r"\bxero\b",
            r"\bmyob\b",
            r"\bquickbooks\b",
            r"\berp\b",
            r"\bsap\b",
            r"\bsystems implementation\b",
            r"\bfinance systems?\b",
            r"\binternal accountant\b",
            r"\btech[- ]savvy accountant\b",
        ]
        if re.search(pattern, text)
    )
    reporting_hits = sum(
        1
        for pattern in [
            r"\bfinancial reporting\b",
            r"\bfinancial statements?\b",
            r"\breconcil",
            r"\bmonth[- ]end\b",
            r"\btax\b",
            r"\bcompliance\b",
            r"\bforecast",
            r"\bgraduate accountant\b",
            r"\bjunior accountant\b",
            r"\bundergraduate accountant\b",
        ]
        if re.search(pattern, text)
    )
    if systems_hits >= max(reporting_hits, 2):
        return "accounting_systems"
    return "accounting_reporting"


def _collect_profile_signal_levels(profile: Dict[str, object], key: str) -> Dict[str, str]:
    signals: Dict[str, str] = {}
    for item in profile.get(key, []) or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("signal") or item.get("tool") or "").strip().lower()
        level = str(item.get("evidence", "")).strip().lower()
        if name and level:
            signals[name] = level
    return signals


def _collect_profile_signal_wording(profile: Dict[str, object], key: str) -> Dict[str, str]:
    wording_map: Dict[str, str] = {}
    for item in profile.get(key, []) or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("signal") or item.get("tool") or "").strip().lower()
        wording = str(item.get("wording", "")).strip()
        if name and wording:
            wording_map[name] = wording
    return wording_map


def classify_finance_signal_evidence(
    profile: Dict[str, object],
    verified_facts: Sequence[str],
    jd_signals: JDSignals,
) -> List[FinanceSignalEvidence]:
    if not is_finance_target(jd_signals):
        return []

    profile_levels = _collect_profile_signal_levels(profile, "finance_capabilities")
    profile_wordings = _collect_profile_signal_wording(profile, "finance_capabilities")
    software_levels = _collect_profile_signal_levels(profile, "accounting_software_exposure")
    software_wordings = _collect_profile_signal_wording(profile, "accounting_software_exposure")
    combined_text = " ".join(str(x) for x in verified_facts).lower()
    evidence: List[FinanceSignalEvidence] = []

    for signal, cfg in FINANCE_SIGNAL_CONFIG.items():
        level = profile_levels.get(signal, "")
        hits = sum(1 for pattern in cfg["patterns"] if re.search(pattern, combined_text))
        if not level:
            if hits >= 2:
                level = "strong"
            elif hits == 1:
                level = "moderate"
            else:
                level = "weak"
        if level not in {"strong", "moderate"}:
            continue
        wording = profile_wordings.get(signal) or (cfg["strong_wording"] if level == "strong" else cfg["moderate_wording"])
        evidence.append(
            FinanceSignalEvidence(
                signal=signal,
                level=level,
                wording=str(wording),
                skill_label=str(cfg["skill_label"]),
                summary_fragment=str(cfg["summary_fragment"]),
                priority=int(cfg["priority"]),
            )
        )

    for tool, cfg in SOFTWARE_CONFIG.items():
        level = software_levels.get(tool, "")
        if level not in {"strong", "moderate"}:
            continue
        wording = software_wordings.get(tool)
        if not wording:
            wording = f"{tool.title()} exposure for accounting system workflows." if tool not in {"sap"} else "Experience with SAP / ERP-supported finance workflows."
        evidence.append(
            FinanceSignalEvidence(
                signal=tool,
                level=level,
                wording=wording,
                skill_label=str(cfg["skill_label"]),
                summary_fragment=str(cfg["summary_fragment"]),
                priority=20 + len(evidence),
            )
        )

    deduped: Dict[str, FinanceSignalEvidence] = {}
    for item in sorted(evidence, key=lambda current: (current.priority, current.signal)):
        deduped.setdefault(item.signal, item)
    return list(deduped.values())


def inject_finance_baseline_signals(
    selected_profile_facts: List[str],
    profile: Dict[str, object],
    verified_facts: Sequence[str],
    jd_signals: JDSignals,
) -> List[FinanceSignalEvidence]:
    finance_signals = classify_finance_signal_evidence(profile, verified_facts, jd_signals)
    existing = {str(fact).strip().lower() for fact in selected_profile_facts}
    for item in finance_signals:
        wording = item.wording.strip()
        if wording.lower() not in existing:
            selected_profile_facts.insert(0, wording)
            existing.add(wording.lower())
    return finance_signals


def finance_signal_prompt_facts(finance_signals: Sequence[FinanceSignalEvidence], max_items: int = 6) -> List[str]:
    return [item.wording for item in sorted(finance_signals, key=lambda current: current.priority)[:max_items] if item.wording.strip()]
