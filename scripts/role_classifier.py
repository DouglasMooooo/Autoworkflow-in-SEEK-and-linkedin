from __future__ import annotations

import re
from typing import Dict, List


ROLE_RULES: Dict[str, Dict[str, List[str]]] = {
    "accounting_reporting": {
        "title": [
            r"\baccountant\b",
            r"\baccounting\b",
            r"\bfinance\b",
            r"\btax\b",
            r"\baudit(or)?\b",
            r"\bproject accountant\b",
        ],
        "jd": [
            r"\bfinancial reporting\b",
            r"\bgeneral ledger\b",
            r"\breconcil",
            r"\bmonth[- ]end\b",
            r"\btax\b",
            r"\bcompliance\b",
            r"\bcash flow\b",
            r"\bfinancial statements?\b",
            r"\bforecast",
            r"\bexcel\b",
        ],
    },
    "accounting_systems": {
        "title": [
            r"\binternal accountant\b",
            r"\bsystems accountant\b",
            r"\bfinance operations\b",
            r"\baccounts officer\b",
            r"\baccounts assistant\b",
            r"\bfinance officer\b",
            r"\baccountant\b",
        ],
        "jd": [
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
            r"\baccounting systems?\b",
            r"\bsystems implementation\b",
            r"\bfinance systems?\b",
            r"\btech[- ]savvy accountant\b",
        ],
    },
    "data_analyst": {
        "title": [
            r"\bdata analyst\b",
            r"\bbi analyst\b",
            r"\breporting analyst\b",
            r"\binsights analyst\b",
        ],
        "jd": [
            r"\bpower bi\b",
            r"\bsql\b",
            r"\btableau\b",
            r"\bdashboard",
            r"\bvisuali[sz]ation\b",
            r"\bdata analysis\b",
            r"\btrend analysis\b",
            r"\bdata quality\b",
            r"\breporting\b",
        ],
    },
    "business_analyst": {
        "title": [
            r"\bbusiness analyst\b",
            r"\bprocess analyst\b",
            r"\boperations analyst\b",
            r"\bbusiness support officer\b",
            r"\bproject support\b",
        ],
        "jd": [
            r"\bbusiness process",
            r"\brequirements\b",
            r"\bstakeholder\b",
            r"\bworkflow",
            r"\bdocumentation\b",
            r"\bprocess improvement\b",
            r"\boperational support\b",
            r"\bcoordination\b",
            r"\breporting\b",
        ],
    },
    "business_graduate": {
        "title": [
            r"\bgraduate\b",
            r"\bgraduate program\b",
            r"\bentry role\b",
            r"\bjunior\b",
        ],
        "jd": [
            r"\brotation",
            r"\btraining\b",
            r"\bcoordination\b",
            r"\bcustomer\b",
            r"\bcommercial\b",
            r"\boperations\b",
            r"\badapt",
            r"\blearn\b",
            r"\breporting\b",
        ],
    },
    "consulting": {
        "title": [r"\bconsulting\b", r"\bconsultant\b", r"\badvisory\b"],
        "jd": [r"\bproblem solving\b", r"\brecommend", r"\bworkshop\b", r"\bstakeholder\b"],
    },
    "sales_bd": {
        "title": [r"\bsales\b", r"\bbusiness development\b", r"\baccount manager\b"],
        "jd": [r"\bclient\b", r"\brevenue\b", r"\bpipeline\b", r"\bnegotiat", r"\bgrowth\b"],
    },
    "operations": {
        "title": [r"\boperations\b", r"\badministration\b", r"\bcoordinator\b"],
        "jd": [r"\bprocess\b", r"\bservice\b", r"\bcoordination\b", r"\bdocumentation\b"],
    },
}


def classify_role_family(job_title: str, cleaned_jd_text: str) -> str:
    text = f"{job_title}\n{cleaned_jd_text}".lower()
    scores = {role: 0.0 for role in ROLE_RULES}

    for role, config in ROLE_RULES.items():
        for pattern in config.get("title", []):
            if re.search(pattern, job_title.lower()):
                scores[role] += 3.2
        for pattern in config.get("jd", []):
            if re.search(pattern, text):
                scores[role] += 1.2

    if re.search(r"\bbusiness analyst\b", job_title.lower()) and scores["business_analyst"] >= scores["data_analyst"]:
        scores["business_analyst"] += 1.4
    if re.search(r"\bdata analyst\b", job_title.lower()):
        scores["data_analyst"] += 1.8
    if re.search(r"\bgraduate\b", job_title.lower()) and not re.search(r"\baccount", job_title.lower()):
        scores["business_graduate"] += 1.6
    if re.search(r"\b(accounts payable|accounts receivable|invoic|billing|quickbooks|xero|myob|erp|sap|systems accountant|internal accountant|finance operations)\b", text):
        scores["accounting_systems"] += 2.8
    if re.search(r"\b(financial reporting|financial statements|tax|compliance|reconcil|month[- ]end|forecast|graduate accountant|junior accountant|undergraduate accountant)\b", text):
        scores["accounting_reporting"] += 2.2
    if re.search(r"\b(public practice|taxation firm|graduate accountant|junior accountant|undergraduate accountant)\b", text):
        scores["accounting_reporting"] += 1.8

    best_role, best_score = max(scores.items(), key=lambda item: item[1])
    return best_role if best_score > 0 else "operations"


def is_accounting_family(role_family: str) -> bool:
    return role_family in {"accounting_reporting", "accounting_systems"}


def define_target_profile(role_family: str, jd_keywords: List[str]) -> List[str]:
    keyword_line = ", ".join(jd_keywords[:5]) if jd_keywords else ""
    if role_family == "accounting_reporting":
        return [
            "Early-career accounting candidate with hands-on experience in tax support, reconciliations, reporting, and compliance-led finance work.",
            "Best evidence comes from accurate filings, financial controls, reporting efficiency, and measurable forecasting outcomes.",
            f"Targeting accounting roles that value Excel, documentation discipline, client support, and practical finance execution{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    if role_family == "accounting_systems":
        return [
            "Finance operations candidate with practical exposure to AP/AR workflows, reconciliations, month-end support, and accounting systems.",
            "Best evidence comes from finance process accuracy, reporting support, ERP exposure, and measurable operational outcomes.",
            f"Targeting systems-oriented accounting roles that value invoicing, reconciliations, accounting software, and dependable finance execution{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    if role_family == "data_analyst":
        return [
            "Data-focused analyst with experience translating raw business data into dashboards, reporting, and decision-ready insight.",
            "Strongest evidence comes from Power BI, Excel, SQL, forecasting, trend analysis, and measurable reporting outcomes.",
            f"Targeting data analyst roles that need clean reporting, dashboard delivery, stakeholder-facing insights, and process support{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    if role_family == "business_analyst":
        return [
            "Business analyst profile with practical exposure to process improvement, operational reporting, stakeholder support, and cross-functional coordination.",
            "Strongest evidence comes from turning messy operational data into clearer workflows, reporting outputs, and business decisions.",
            f"Targeting business analyst roles that value requirements understanding, documentation, workflow support, and business problem solving{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    if role_family == "business_graduate":
        return [
            "Early-career business graduate profile with commercial exposure, reporting capability, adaptability, and strong learning velocity.",
            "Strongest evidence comes from measurable business outcomes, customer-facing execution, and cross-functional support under pressure.",
            f"Targeting graduate roles that combine operational support, reporting, stakeholder communication, and growth potential{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    if role_family == "consulting":
        return [
            "Structured problem solver with stakeholder-facing communication and analytical rigor.",
            "Strongest evidence comes from synthesising data, clarifying business issues, and recommending practical actions.",
            f"Targeting consulting-style work that values ownership, business judgment, and communication{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    if role_family == "sales_bd":
        return [
            "Commercially minded profile with client-facing communication, service responsiveness, and growth orientation.",
            "Strongest evidence comes from measurable revenue support, relationship management, and execution in fast-paced commercial settings.",
            f"Targeting sales and BD roles that value customer impact, persistence, and commercial awareness{f' across {keyword_line}' if keyword_line else ''}.",
        ]
    return [
        "Execution-focused profile with practical cross-functional experience and strong operating discipline.",
        "Strongest evidence comes from process support, reporting, coordination, and reliable day-to-day delivery.",
        f"Targeting roles that need dependable execution, communication, and structured support{f' across {keyword_line}' if keyword_line else ''}.",
    ]
