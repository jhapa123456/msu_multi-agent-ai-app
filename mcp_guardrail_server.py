"""
MCP-style Guardrail Server for the MSU Catalog Agentic RAG demo.

This is intentionally lightweight so it deploys on Streamlit Community Cloud without
running a separate long-lived server. In production, these functions can be exposed
as tools through a real MCP server and called by LangGraph/LangChain agents.
"""
from __future__ import annotations

import re
from urllib.parse import urlparse
from typing import Dict, List

ALLOWED_DOMAINS = {"catalog.msutexas.edu", "msutexas.edu", "www.msutexas.edu"}
HIGH_STAKES_TOPICS = [
    "visa", "immigration", "financial aid", "scholarship", "dismissal",
    "appeal", "probation", "residency", "tuition", "deadline", "admission decision",
    "graduation", "degree audit", "transfer credit", "international"
]


def authorize_url(url: str) -> bool:
    """Allow only approved MSU Texas domains."""
    try:
        host = urlparse(url).netloc.lower()
        return any(host == d or host.endswith("." + d) for d in ALLOWED_DOMAINS)
    except Exception:
        return False


def sanitize_question(question: str) -> str:
    """Remove obvious prompt-injection strings while keeping the student question usable."""
    question = question.strip()
    bad_patterns = [
        r"ignore previous instructions",
        r"reveal system prompt",
        r"developer message",
        r"tool output",
        r"jailbreak",
        r"act as dan",
    ]
    for pat in bad_patterns:
        question = re.sub(pat, "[removed]", question, flags=re.IGNORECASE)
    return question[:1000]


def detect_pii(text: str) -> List[str]:
    findings = []
    if re.search(r"\b\d{3}-\d{2}-\d{4}\b", text):
        findings.append("possible_ssn")
    if re.search(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", text, re.I):
        findings.append("email")
    if re.search(r"\b\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b", text):
        findings.append("phone")
    return findings


def is_high_stakes(question: str) -> bool:
    q = question.lower()
    return any(topic in q for topic in HIGH_STAKES_TOPICS)


def validate_answer(answer: str, citations: List[Dict]) -> Dict:
    """Check that final answer has citations and does not overclaim."""
    issues = []
    if not citations:
        issues.append("No source citation attached.")
    if any(word in answer.lower() for word in ["guaranteed", "definitely admitted", "100%"]):
        issues.append("Potentially overconfident policy claim.")
    if detect_pii(answer):
        issues.append("Answer appears to contain personal information.")
    return {
        "allowed": len(issues) == 0,
        "issues": issues,
        "recommendation": "Show advisor escalation note for high-stakes or low-evidence answers."
    }


def guardrail_summary() -> Dict:
    return {
        "name": "MCP-style Guardrail Server",
        "purpose": "Security, source authorization, PII detection, citation validation, and escalation control.",
        "production_upgrade": "Expose these functions as MCP tools behind an authenticated MCP server with RBAC, audit logs, and allowlisted connectors."
    }
