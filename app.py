import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import datetime
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TPCRA v3.0 Dashboard",
    page_icon="🔐",
    layout="wide",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
.kpi-label { font-size: 11px; font-weight: 600; text-transform: uppercase;
             letter-spacing: 0.07em; color: #6c757d; margin-bottom: 4px; }
.kpi-value { font-size: 28px; font-weight: 600; line-height: 1.1; }
.tier-critical { color: #A32D2D; }
.tier-high     { color: #854F0B; }
.tier-medium   { color: #185FA5; }
.tier-low      { color: #3B6D11; }
.response-pill {
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 11px; font-weight: 600; letter-spacing: 0.03em;
}
.stTabs [data-baseweb="tab"] { font-size: 13px; padding: 8px 16px; }
div[data-testid="metric-container"] {
    background: #f8f9fa; border-radius: 10px; padding: 14px;
    border: 1px solid #e9ecef;
}
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
DOMAIN_MAP = {
    "A": "Organizational Management",
    "B": "Human Resource Management",
    "C": "Infrastructure Security",
    "D": "Data Protection",
    "E": "Access Management",
    "F": "Application Security",
    "G": "System Security",
    "H": "Email Security",
    "I": "Mobile Devices",
    "J": "Incident Response",
    "K": "Cloud Services",
    "L": "Business Continuity",
    "M": "Supply Chain & Physical Security",
    "N": "AI & Emerging Technology Risk",
}

RESP_COLORS = {
    "Yes":     "#639922",
    "No":      "#E24B4A",
    "Partial": "#EF9F27",
    "N/A":     "#B4B2A9",
}

RESP_PILL = {
    "Yes":     "background:#EAF3DE;color:#27500A",
    "No":      "background:#FCEBEB;color:#A32D2D",
    "Partial": "background:#FAEEDA;color:#633806",
    "N/A":     "background:#F1EFE8;color:#5F5E5A",
    "—":       "background:#F1EFE8;color:#888780",
}

TIER_PILL = {
    "Critical": "background:#FCEBEB;color:#A32D2D",
    "High":     "background:#FAEEDA;color:#854F0B",
    "Medium":   "background:#E6F1FB;color:#0C447C",
    "Low":      "background:#EAF3DE;color:#3B6D11",
}

RATING_THRESHOLDS = [
    (90, "✅ Low",      "#639922"),
    (70, "🟡 Medium",   "#BA7517"),
    (50, "🟠 High",     "#D85A30"),
    (0,  "🔴 Critical", "#A32D2D"),
]

YES_VALS  = {"yes", "y"}
NO_VALS   = {"no", "n"}
NA_VALS   = {"n/a", "na", "not applicable"}
PART_VALS = {"partial", "partly", "partially"}
EVIDENCE_STATUS = {"submitted", "provided", "received", "complete", "done", "yes"}

# ── Static gap database (from TPCRA v3.0 questionnaire analysis) ───────────────
GAP_DB = [
    {
        "id": "A", "name": "Organizational Management", "risk": "Critical",
        "critical": 3, "high": 4, "medium": 2, "total": 9,
        "gaps": [
            {"ref": "A.1",  "text": "IT Security policies not formally established or documented", "tier": "Critical"},
            {"ref": "A.4e", "text": "Security incident response not covered in policy", "tier": "Critical"},
            {"ref": "A.6",  "text": "Policy not aligned to regulatory requirements / customer data protection", "tier": "Critical"},
            {"ref": "A.5",  "text": "Policies not formally acknowledged / signed by all staff", "tier": "High"},
            {"ref": "A.7",  "text": "No formal risk management framework (ISO 27001, NIST CSF)", "tier": "High"},
            {"ref": "A.8",  "text": "No independent security certification or audit report", "tier": "High"},
        ],
        "recs": [
            "Establish and approve a formal Information Security Policy covering all required topics (A.4a–k) within 30 days",
            "Align policies to applicable regulations (DPA, GDPR, PDPA) and obtain senior leadership sign-off",
            "Initiate ISO 27001 or NIST CSF gap assessment; target certification within 12 months",
            "Obtain and share current third-party security audit report or certification",
        ],
    },
    {
        "id": "B", "name": "Human Resource Management", "risk": "Critical",
        "critical": 2, "high": 2, "medium": 1, "total": 6,
        "gaps": [
            {"ref": "B.3", "text": "Background screening and NDA not enforced for new hires / contractors", "tier": "Critical"},
            {"ref": "B.5", "text": "No formal offboarding process to revoke access on last working day", "tier": "Critical"},
            {"ref": "B.2", "text": "Security awareness training not extended to all employees and contractors", "tier": "High"},
            {"ref": "B.6", "text": "No insider threat controls (UBA, session monitoring, separation of duties)", "tier": "High"},
        ],
        "recs": [
            "Implement mandatory background checks and NDA execution before granting system access",
            "Formalize offboarding checklist to revoke all logical and physical access on day of departure",
            "Roll out annual security awareness training to 100% of staff, including third-party contractors",
            "Deploy user behaviour analytics or privileged session monitoring to detect insider threats",
        ],
    },
    {
        "id": "C", "name": "Infrastructure Security", "risk": "Critical",
        "critical": 11, "high": 7, "medium": 0, "total": 20,
        "gaps": [
            {"ref": "C.1.2", "text": "No clear network segmentation; externally-facing apps not in DMZ", "tier": "Critical"},
            {"ref": "C.1.3", "text": "Corporate network not separated from internet by firewall", "tier": "Critical"},
            {"ref": "C.1.5", "text": "Inter-system communication not using encrypted protocols only", "tier": "Critical"},
            {"ref": "C.1.6", "text": "Deprecated protocols (TLS 1.0/1.1, Telnet, FTP) not disabled", "tier": "Critical"},
            {"ref": "C.2.1", "text": "IPS/firewall not blocking unauthorized external access", "tier": "Critical"},
            {"ref": "C.2.4", "text": "WAF not deployed on internet-facing applications", "tier": "Critical"},
            {"ref": "C.5.1", "text": "Remote connections do not require MFA and encrypted channel", "tier": "Critical"},
            {"ref": "C.5.4", "text": "Remote access not revoked promptly (within 24 hours) when no longer required", "tier": "Critical"},
            {"ref": "C.6.2", "text": "No defined SLA timelines for applying security patches by criticality", "tier": "Critical"},
            {"ref": "C.3.5", "text": "No centralized security event monitoring (SIEM/SOC)", "tier": "Critical"},
        ],
        "recs": [
            "Implement network segmentation with DMZ for all internet-facing services; firewall all public access points",
            "Disable TLS 1.0/1.1, SSLv3, FTP, Telnet, and enforce TLS 1.2+ across all systems immediately",
            "Deploy WAF and enforce MFA for all remote and cloud console access",
            "Define and publish patch SLA: Critical ≤24h, High ≤7 days, Medium ≤30 days",
            "Deploy SIEM or engage SOC for centralized log aggregation and alerting",
        ],
    },
    {
        "id": "D", "name": "Data Protection", "risk": "Critical",
        "critical": 8, "high": 4, "medium": 1, "total": 14,
        "gaps": [
            {"ref": "D.1.3",  "text": "No technical isolation of client data from other tenants", "tier": "Critical"},
            {"ref": "D.1.4",  "text": "Data not protected by encryption and access controls across lifecycle", "tier": "Critical"},
            {"ref": "D.1.5",  "text": "No process for secure data destruction on contract termination", "tier": "Critical"},
            {"ref": "D.1.6",  "text": "Controls missing to prevent unauthorized extraction of customer data", "tier": "Critical"},
            {"ref": "D.2.1",  "text": "Strong cryptographic protocols (TLS 1.2+) not protecting data in transit", "tier": "Critical"},
            {"ref": "D.2.3",  "text": "Cryptographic key management lacks dual-control and secure storage", "tier": "Critical"},
            {"ref": "D.2.4",  "text": "Backup data not encrypted using industry-accepted standards", "tier": "Critical"},
            {"ref": "D.2.5",  "text": "Deprecated cipher suites (MD5, SHA-1, RC4) not identified and disabled", "tier": "Critical"},
        ],
        "recs": [
            "Implement data classification tagging and enforce access controls aligned to sensitivity level",
            "Encrypt all data at rest (AES-256) and in transit (TLS 1.2+); include backup media",
            "Establish a formal key management lifecycle policy with dual-control and secure vaulting",
            "Define and enforce data retention and secure destruction procedures, including for backups",
            "Deploy DLP solution to detect and block unauthorized data exfiltration",
        ],
    },
    {
        "id": "E", "name": "Access Management", "risk": "Critical",
        "critical": 8, "high": 7, "medium": 2, "total": 18,
        "gaps": [
            {"ref": "E.1.1",  "text": "Access to critical systems not gated by formal approved request process", "tier": "Critical"},
            {"ref": "E.1.3",  "text": "No periodic review of user access to systems and databases", "tier": "Critical"},
            {"ref": "E.1.7",  "text": "Shared user IDs or privileged accounts in use", "tier": "Critical"},
            {"ref": "E.1.8",  "text": "MFA not enforced for remote, cloud, and privileged access", "tier": "Critical"},
            {"ref": "E.2.3",  "text": "MFA not enforced for all user accounts", "tier": "Critical"},
            {"ref": "E.2.12", "text": "Passwords not stored with strong salted hashing (bcrypt/Argon2)", "tier": "Critical"},
            {"ref": "E.3.1",  "text": "No documented procedures for privileged / break-glass access", "tier": "Critical"},
            {"ref": "E.3.2",  "text": "Privileged access not reviewed quarterly with dual-control / JIT", "tier": "Critical"},
        ],
        "recs": [
            "Enforce MFA organisation-wide — prioritise privileged, remote, and cloud console access",
            "Implement quarterly user access reviews with automated provisioning/deprovisioning workflows",
            "Eliminate shared accounts; enforce unique user IDs and least-privilege principles",
            "Deploy or evaluate PAM solution for privileged credential vaulting and session recording",
            "Migrate to bcrypt or Argon2 password hashing; audit credential storage immediately",
        ],
    },
    {
        "id": "F", "name": "Application Security", "risk": "Critical",
        "critical": 4, "high": 4, "medium": 1, "total": 10,
        "gaps": [
            {"ref": "F.3", "text": "No security / pentest conducted before application onboarding", "tier": "Critical"},
            {"ref": "F.4", "text": "Secure coding guidelines (OWASP Top 10) not formally adopted", "tier": "Critical"},
            {"ref": "F.9", "text": "API security controls (OAuth, rate-limiting, input validation) not implemented", "tier": "Critical"},
            {"ref": "F.6", "text": "Security requirements not formally embedded in SDLC", "tier": "High"},
            {"ref": "F.7", "text": "No SCA process for open-source / third-party library vulnerabilities", "tier": "High"},
        ],
        "recs": [
            "Mandate penetration testing and secure code review before any application goes live",
            "Adopt OWASP Top 10 and SANS CWE Top 25 as formal secure coding standards",
            "Implement API gateway with OAuth 2.0, rate limiting, and input validation",
            "Integrate SAST/DAST tooling and SCA into CI/CD pipelines",
        ],
    },
    {
        "id": "G", "name": "System Security", "risk": "Critical",
        "critical": 6, "high": 4, "medium": 0, "total": 11,
        "gaps": [
            {"ref": "G.2.1", "text": "No timely review of security audit logs; anomalies not acted on", "tier": "Critical"},
            {"ref": "G.2.2", "text": "Automated audit trails not capturing admin actions and auth changes", "tier": "Critical"},
            {"ref": "G.2.3", "text": "Security logs not reviewed for anomalies via SIEM correlation", "tier": "Critical"},
            {"ref": "G.2.4", "text": "Logging facilities not protected against tampering / unauthorized modification", "tier": "Critical"},
            {"ref": "G.3.1", "text": "No real-time monitoring of security events on critical systems", "tier": "Critical"},
            {"ref": "G.3.5", "text": "No quarterly automated vulnerability scanning with defined SLA remediation", "tier": "Critical"},
        ],
        "recs": [
            "Deploy SIEM with automated correlation rules; assign SOC team / alerting coverage 24×7",
            "Implement tamper-protected, write-once log storage with minimum 12-month retention",
            "Schedule quarterly automated vulnerability scans with SLA-tracked remediation",
            "Enable full audit trails for admin actions, authentication events, and access changes",
        ],
    },
    {
        "id": "H", "name": "Email Security", "risk": "High",
        "critical": 1, "high": 4, "medium": 0, "total": 5,
        "gaps": [
            {"ref": "H.3", "text": "Email attachments and URLs not scanned for malware/phishing (sandboxing)", "tier": "Critical"},
            {"ref": "H.5", "text": "SPF, DKIM, DMARC email authentication protocols not implemented", "tier": "High"},
            {"ref": "H.2", "text": "Email data confidentiality and integrity not protected by technical controls", "tier": "High"},
            {"ref": "H.6", "text": "Simulated phishing exercises not conducted at least annually", "tier": "High"},
        ],
        "recs": [
            "Deploy email gateway with sandboxing for URL and attachment analysis",
            "Implement SPF, DKIM, and DMARC with at minimum quarantine policy",
            "Conduct simulated phishing exercises at least annually; track click rates",
        ],
    },
    {
        "id": "I", "name": "Mobile Devices", "risk": "High",
        "critical": 1, "high": 4, "medium": 0, "total": 5,
        "gaps": [
            {"ref": "I.3", "text": "Laptops do not enforce full-disk encryption (BitLocker / FileVault)", "tier": "Critical"},
            {"ref": "I.2", "text": "Mobile devices not enrolled in MDM with security software / PIN enforced", "tier": "High"},
            {"ref": "I.5", "text": "BYOD devices lack corporate data isolation and remote wipe capability", "tier": "High"},
            {"ref": "I.6", "text": "Jailbroken/rooted devices not automatically blocked from corporate access", "tier": "High"},
        ],
        "recs": [
            "Enforce full-disk encryption on all laptops via Group Policy or MDM",
            "Enrol all corporate and BYOD devices into MDM; enforce PIN, remote wipe, and compliance check",
            "Block jailbroken/rooted devices via MDM compliance policy",
        ],
    },
    {
        "id": "J", "name": "Incident Response", "risk": "Critical",
        "critical": 3, "high": 3, "medium": 1, "total": 7,
        "gaps": [
            {"ref": "J.2", "text": "No documented Incident Response Plan with roles and escalation paths", "tier": "Critical"},
            {"ref": "J.3", "text": "IRP not tested at least annually (tabletop / simulation)", "tier": "Critical"},
            {"ref": "J.5", "text": "No defined contractual SLA for notifying client of security incidents", "tier": "Critical"},
            {"ref": "J.4", "text": "IRP not reviewed and updated at least annually", "tier": "High"},
            {"ref": "J.7", "text": "No 24×7 security operations capability (in-house or outsourced SOC)", "tier": "High"},
        ],
        "recs": [
            "Develop and approve an Incident Response Plan covering all incident types with escalation matrix",
            "Conduct annual tabletop exercise; document and track lessons learned",
            "Define and agree contractual incident notification SLA (recommend ≤4 hours for confirmed breaches)",
            "Establish or contract 24×7 SOC capability for continuous threat monitoring",
        ],
    },
    {
        "id": "K", "name": "Cloud Services", "risk": "Critical",
        "critical": 7, "high": 3, "medium": 1, "total": 11,
        "gaps": [
            {"ref": "K.2", "text": "CSP lacks third-party security certification (ISO 27001, SOC 2, etc.)", "tier": "Critical"},
            {"ref": "K.4", "text": "No logical or physical data segregation between tenants", "tier": "Critical"},
            {"ref": "K.5", "text": "Hypervisor admin access not restricted with least privilege and MFA", "tier": "Critical"},
            {"ref": "K.6", "text": "Cryptographic key management process not formal (AWS KMS / Azure Key Vault)", "tier": "Critical"},
            {"ref": "K.8", "text": "Security logs not generated, retained, and accessible on demand", "tier": "Critical"},
            {"ref": "K.9", "text": "No confirmed data deletion process with certificate upon termination", "tier": "Critical"},
        ],
        "recs": [
            "Verify CSP holds current ISO 27001, SOC 2 Type 2, or equivalent certification",
            "Confirm tenant data isolation and request Shared Responsibility Matrix from CSP",
            "Enforce MFA and least privilege for all hypervisor and admin console access",
            "Implement formal key management via AWS KMS / Azure Key Vault; document lifecycle",
            "Obtain written data deletion certificate commitment; include in contract",
        ],
    },
    {
        "id": "L", "name": "Business Continuity", "risk": "Critical",
        "critical": 4, "high": 2, "medium": 0, "total": 6,
        "gaps": [
            {"ref": "L.1", "text": "No formally approved Business Continuity Plan (BCP)", "tier": "Critical"},
            {"ref": "L.3", "text": "BCP not tested at least annually", "tier": "Critical"},
            {"ref": "L.4", "text": "RTO and RPO targets not defined or validated through testing", "tier": "Critical"},
            {"ref": "L.5", "text": "Backups not stored off-site or tested for restorability annually", "tier": "Critical"},
        ],
        "recs": [
            "Develop, approve, and communicate a formal BCP with RTO/RPO targets for critical services",
            "Schedule annual BCP/DR tabletop and failover test; document results and remediation actions",
            "Implement off-site or cross-region encrypted backup storage with regular restore tests",
        ],
    },
    {
        "id": "M", "name": "Supply Chain & Physical Security", "risk": "Critical",
        "critical": 3, "high": 5, "medium": 2, "total": 10,
        "gaps": [
            {"ref": "M.1", "text": "No formal TPRM program to assess sub-processors", "tier": "Critical"},
            {"ref": "M.2", "text": "Sub-processors / fourth parties with data access not identified", "tier": "Critical"},
            {"ref": "M.3", "text": "Security requirements not contractually flowed down to sub-processors", "tier": "Critical"},
            {"ref": "M.6", "text": "Physical access to server rooms not restricted with multi-factor controls", "tier": "High"},
            {"ref": "M.9", "text": "Environmental controls (fire suppression, UPS, CCTV) not confirmed for data hosting", "tier": "High"},
        ],
        "recs": [
            "Establish a TPRM program and complete security assessments for all sub-processors",
            "Inventory and disclose all fourth parties with access to client data",
            "Include equivalent security obligations in all sub-processor contracts",
            "Confirm physical access controls and environmental safeguards for data facilities",
        ],
    },
    {
        "id": "N", "name": "AI & Emerging Technology Risk", "risk": "Critical",
        "critical": 9, "high": 4, "medium": 0, "total": 14,
        "gaps": [
            {"ref": "N.1.1", "text": "No formal AI usage policy approved by senior leadership", "tier": "Critical"},
            {"ref": "N.1.2", "text": "No AI risk assessment; roles and responsibilities for AI governance undefined", "tier": "Critical"},
            {"ref": "N.2.1", "text": "Client data may be used to train AI models without explicit consent", "tier": "Critical"},
            {"ref": "N.2.2", "text": "No technical controls preventing PII from entering AI models", "tier": "Critical"},
            {"ref": "N.2.3", "text": "AI outputs not validated before use in customer-affecting decisions", "tier": "Critical"},
            {"ref": "N.3.1", "text": "Third-party AI services lack DPA confirming no training on client data", "tier": "Critical"},
            {"ref": "N.4.1", "text": "AI systems lack RBAC and MFA; not hardened at deployment", "tier": "Critical"},
            {"ref": "N.4.2", "text": "No controls to detect prompt injection or adversarial input attacks", "tier": "Critical"},
            {"ref": "N.5.1", "text": "No tamper-protected audit logs for AI interactions", "tier": "Critical"},
        ],
        "recs": [
            "Publish an AI usage policy (acceptable use, prohibited inputs, output handling) approved by CISO/DPO",
            "Conduct a formal AI risk assessment; define AI governance roles and incident response procedures",
            "Require DPA from all third-party AI providers confirming client data is not used for model training",
            "Implement PII masking/tokenization before data reaches AI models; enable prompt/output logging",
            "Deploy prompt injection detection and output guardrails; enforce RBAC + MFA on all AI system access",
            "Maintain tamper-protected audit logs for all AI interactions with regular anomaly review",
        ],
    },
]


# ── Helpers ────────────────────────────────────────────────────────────────────
def normalize_response(val) -> str:
    if val is None:
        return "—"
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%m/%d/%Y")
    s = str(val).strip().lower()
    if s in YES_VALS:  return "Yes"
    if s in NO_VALS:   return "No"
    if s in NA_VALS:   return "N/A"
    if s in PART_VALS: return "Partial"
    if not s or s == "—": return "—"
    return str(val).strip()


def extract_domain(key) -> str:
    if not key:
        return ""
    s = str(key).strip()
    # "A — ORGANIZATIONAL MANAGEMENT" style headers
    if " — " in s:
        return s.split(" — ")[0].strip().upper()
    if s and s[0].isalpha():
        return s[0].upper()
    return ""


def compliance_score(items: list) -> int:
    scored = [i for i in items if i["norm"] not in ("—", "N/A")]
    if not scored:
        return 0
    earned = sum(100 if i["norm"] == "Yes" else (50 if i["norm"] == "Partial" else 0)
                 for i in scored)
    return round(earned / len(scored))


def risk_rating(score: int) -> tuple:
    for threshold, label, color in RATING_THRESHOLDS:
        if score >= threshold:
            return label, color
    return "🔴 Critical", "#A32D2D"


def pill(text: str, style: str) -> str:
    return f'<span class="response-pill" style="{style}">{text}</span>'


def resp_pill(norm: str) -> str:
    return pill(norm if norm != "—" else "—", RESP_PILL.get(norm, RESP_PILL["—"]))


def tier_pill(tier: str) -> str:
    return pill(tier, TIER_PILL.get(tier, "background:#F1EFE8;color:#888780"))


# ── Parsers ────────────────────────────────────────────────────────────────────
def parse_part1(wb) -> dict:
    """Parse Part 1 — Contact & Engagement information."""
    if "Part 1" not in wb.sheetnames:
        return {}
    ws = wb["Part 1"]
    rows = list(ws.iter_rows(values_only=True))
    meta = {"title": "", "sections": {}, "items": []}

    current_section = ""
    for row in rows:
        if not any(v is not None for v in row):
            continue
        key, question, response = row[0], row[1], row[2] if len(row) > 2 else None

        # Title row
        if isinstance(key, str) and "TPCRA" in str(key) and not meta["title"]:
            meta["title"] = str(key).strip()
            continue

        # Section header
        if isinstance(key, str) and key.startswith("SECTION"):
            current_section = str(key).strip()
            meta["sections"].setdefault(current_section, [])
            continue

        # Data row
        if question and str(question).strip():
            item = {
                "key":      str(key).strip() if key else "",
                "section":  current_section,
                "question": str(question).strip(),
                "response": str(response).strip() if response else "",
                "other":    str(row[3]).strip() if len(row) > 3 and row[3] else "",
                "tier":     str(row[4]).strip() if len(row) > 4 and row[4] and str(row[4]) != "—" else "",
            }
            meta["items"].append(item)
            if current_section:
                meta["sections"][current_section].append(item)

    return meta


def parse_part2(wb) -> dict:
    """Parse Part 2 — Security questionnaire with responses, tiers, and remarks."""
    if "Part 2" not in wb.sheetnames:
        return {}
    ws = wb["Part 2"]
    rows = list(ws.iter_rows(values_only=True))

    result = {"title": "", "domains": {}, "items": []}
    current_domain = ""
    current_sub = ""

    for row in rows:
        if not any(v is not None for v in row):
            continue
        key = row[0]
        question  = row[1] if len(row) > 1 else None
        response  = row[2] if len(row) > 2 else None
        other     = row[3] if len(row) > 3 else None
        tier      = row[4] if len(row) > 4 else None

        key_s = str(key).strip() if key else ""

        # Title row
        if isinstance(key, str) and "TPCRA" in key and not result["title"]:
            result["title"] = key.strip()
            continue

        # Column header row
        if key_s == "#":
            continue

        # Domain header: "A — ORGANIZATIONAL MANAGEMENT"
        if isinstance(key, str) and " — " in key and key[0].isalpha() and len(key.split(" — ")[0]) == 1:
            letter = key.split(" — ")[0].strip().upper()
            name   = key.split(" — ", 1)[1].strip()
            current_domain = letter
            current_sub = ""
            result["domains"].setdefault(letter, {
                "name": DOMAIN_MAP.get(letter, name),
                "items": []
            })
            continue

        # Sub-section label rows (no response, key has a dot or is descriptive text)
        if isinstance(key, str) and question and response is None and tier is None:
            current_sub = str(question).strip() if question else ""
            continue

        # Question row — must have a proper key like A.1, B.2.3, N.1.1 etc.
        if key_s and question and str(question).strip():
            # Determine domain from key
            domain_letter = key_s[0].upper() if key_s[0].isalpha() else current_domain

            tier_s = str(tier).strip() if tier and str(tier).strip() not in ("—", "None", "") else ""
            norm   = normalize_response(response)

            item = {
                "key":      key_s,
                "domain":   domain_letter,
                "domain_name": DOMAIN_MAP.get(domain_letter, domain_letter),
                "sub":      current_sub,
                "question": str(question).strip(),
                "response": str(response).strip() if response else "",
                "norm":     norm,
                "other":    str(other).strip() if other else "",
                "tier":     tier_s,
            }

            result["items"].append(item)
            if domain_letter in result["domains"]:
                result["domains"][domain_letter]["items"].append(item)
            elif domain_letter:
                result["domains"].setdefault(domain_letter, {
                    "name": DOMAIN_MAP.get(domain_letter, domain_letter),
                    "items": [item]
                })

    return result


def parse_evidence(wb) -> list:
    """Parse Evidence checklist sheet."""
    if "Evidence" not in wb.sheetnames:
        return []
    ws = wb["Evidence"]
    rows = list(ws.iter_rows(values_only=True))
    items = []
    for row in rows[2:]:
        if not any(v is not None for v in row):
            continue
        num, evidence, guidance, status, remarks, required_for = (
            row[0], row[1], row[2], row[3], row[4], row[5] if len(row) > 5 else None
        )
        if evidence:
            status_norm = "Submitted" if str(status or "").strip().lower() in EVIDENCE_STATUS else (
                str(status).strip() if status else "Pending"
            )
            items.append({
                "num":          str(num).strip() if num else "",
                "evidence":     str(evidence).strip(),
                "guidance":     str(guidance).strip() if guidance else "",
                "status":       status_norm,
                "remarks":      str(remarks).strip() if remarks else "",
                "required_for": str(required_for).strip() if required_for else "",
            })
    return items


def extract_contact(p1_items: list) -> dict:
    """Pull key contact fields from Part 1 items."""
    contact = {"vendor": "", "rep": "", "email": "", "engagement": ""}
    for item in p1_items:
        q = item["question"].lower()
        r = item["response"]
        if not r or r in ("None", "—"):
            continue
        if "company name" in q:
            contact["vendor"] = r
        elif "authorized representative" in q and "email" not in q:
            contact["rep"] = r
        elif "email" in q and "representative" in q:
            contact["email"] = r
        elif "description of the engagement" in q:
            contact["engagement"] = r[:120] + "…" if len(r) > 120 else r
    return contact


def make_sample_excel() -> bytes:
    """Generate a minimal sample Excel template matching TPCRA v3.0 format."""
    p2_rows = [
        ("TPCRA Questionnaire - Part 2  |  v3.0  |  Response options: Yes / No / Partial / N/A", None, None, None, None, None),
        ("#", "Statement / Question", "Response\n(Yes/No/Partial/N/A)", "Other Information\n(Remarks & Evidence)", "Risk\nTier", "Comments\nRequired"),
        ("A — ORGANIZATIONAL MANAGEMENT", None, None, None, None, None),
        ("A.1", "IT Security policies and procedures are formally established and documented.", None, None, "Critical", "—"),
        ("A.2", "IT Security policies and procedures are reviewed at least annually.", None, None, "High", "—"),
        ("A.5", "IT Security policies are formally acknowledged by all employees and contractors.", None, None, "High", "—"),
        ("A.6", "IT Security policies comply with relevant regulatory requirements.", None, None, "Critical", "—"),
        ("B — HUMAN RESOURCE MANAGEMENT", None, None, None, None, None),
        ("B.2", "IT security awareness training is provided to ALL employees.", None, None, "High", "—"),
        ("B.3", "New hires are subjected to background screening and must sign an NDA.", None, None, "Critical", "—"),
        ("B.5", "A formal employee offboarding process revokes all access on the last working day.", None, None, "Critical", "—"),
        ("C — INFRASTRUCTURE SECURITY", None, None, None, None, None),
        ("C.1.2", "Clear network segmentation exists between web, application, and database tiers.", None, None, "Critical", "—"),
        ("C.1.5", "All inter-system communication uses encrypted protocols only (HTTPS, SFTP, TLS 1.2+).", None, None, "Critical", "—"),
        ("C.5.1", "All remote connections require MFA and encrypted communications.", None, None, "Critical", "—"),
        ("C.6.2", "Security patches are applied within defined SLA timelines based on criticality.", None, None, "Critical", "—"),
        ("C.7.2", "Anti-malware solutions are deployed on all user computers, servers, and endpoints.", None, None, "Critical", "—"),
        ("D — DATA PROTECTION", None, None, None, None, None),
        ("D.1.3", "Technical measures isolate company data from other clients.", None, None, "Critical", "—"),
        ("D.1.10", "A DLP solution or equivalent is deployed to prevent unauthorized data exfiltration.", None, None, "High", "—"),
        ("D.2.1", "Strong cryptographic protocols protect sensitive data in transit.", None, None, "Critical", "—"),
        ("E — ACCESS MANAGEMENT", None, None, None, None, None),
        ("E.1.1", "Access to critical systems is granted only on a need-to-know basis.", None, None, "Critical", "—"),
        ("E.1.7", "Sharing of user IDs or privileged accounts is strictly prohibited.", None, None, "Critical", "—"),
        ("E.1.8", "MFA is enforced for all remote access and privileged account usage.", None, None, "Critical", "—"),
        ("J — INCIDENT RESPONSE", None, None, None, None, None),
        ("J.2", "A documented Incident Response Plan (IRP) is established.", None, None, "Critical", "—"),
        ("J.3", "The IRP is tested at least annually.", None, None, "Critical", "—"),
        ("N — AI & EMERGING TECHNOLOGY RISK", None, None, None, None, None),
        ("N.1.1", "A formal AI usage policy is in place covering acceptable use and accountability.", None, None, "Critical", "—"),
        ("N.2.1", "Data will NOT be used to train AI models without explicit written consent.", None, None, "Critical", "—"),
    ]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df = pd.DataFrame(p2_rows)
        df.to_excel(writer, index=False, header=False, sheet_name="Part 2")
    return buf.getvalue()


# ── PDF generation ────────────────────────────────────────────────────────────
def generate_vendor_pdf(gap_db: list, vendor_name: str = "", assessment_date: str = "") -> bytes:
    """
    Generate a vendor-facing gap summary PDF (2 pages) using ReportLab.
    Page 1 — Gap overview: KPIs + domain gap table + expectations note.
    Page 2 — NIST-grounded recommendations per domain.
    """

    # ── NIST-grounded recommendations database ─────────────────────────────────
    # Sources: NIST SP 800-53r5, NIST SP 800-161r1, NIST CSF 2.0
    NIST_RECS = {
        "A": {
            "title": "Organizational Management",
            "nist_refs": "NIST CSF 2.0: GV.PO / NIST SP 800-53r5: PL-1, PL-2, PM-1, PM-9",
            "recs": [
                ("Establish & document an IS policy", "Formally document and approve an Information Security Policy covering all required topics. Per NIST SP 800-53r5 PL-1, policies must state purpose, scope, roles, responsibilities, and management commitment, and be reviewed at defined intervals."),
                ("Adopt a risk management framework", "Implement a formal risk management framework such as NIST CSF 2.0 or ISO 27001. NIST SP 800-161r1 requires enterprises to integrate C-SCRM into enterprise-wide risk management and establish an ERM council with CISO participation (SP 800-161r1 Sec. 2)."),
                ("Obtain independent certification", "Pursue and maintain a third-party security certification (ISO 27001, SOC 2 Type 2). NIST SP 800-161r1 SA-9 requires periodic revalidation of supplier adherence to security requirements via requisite certifications, site visits, or third-party assessments."),
                ("Enforce policy acknowledgement", "Ensure all staff and contractors formally acknowledge security policies. NIST SP 800-53r5 PS-6 requires signed access agreements that reference policy obligations for all personnel, including contractor staff."),
            ],
        },
        "B": {
            "title": "Human Resource Management",
            "nist_refs": "NIST SP 800-53r5: PS-3, PS-6, PS-7, AT-2, AT-3 / NIST SP 800-161r1: PS-3, PS-7",
            "recs": [
                ("Screen all personnel before access", "Extend background screening to contractors and third-party personnel. Per NIST SP 800-161r1 PS-3, personnel screening policies must apply to any contractor with authorized system access, with continuous monitoring commensurate with access level, and flowed down to sub-tier contractors."),
                ("Formalize offboarding access revocation", "Implement a formal offboarding process that revokes all logical and physical access on the last day. NIST SP 800-53r5 AC-3(8) requires prompt revocation of access authorizations for contractors who no longer require access, including retirement of credentials and disabling of all accounts."),
                ("Deliver role-based security training", "Provide security awareness training to all staff including contractors. NIST SP 800-53r5 AT-2 and AT-3 require role-based training covering insider threat (AT-2(2)), social engineering (AT-2(3)), and phishing (AT-2(4)), with completion records maintained (AT-4)."),
                ("Deploy insider threat controls", "Implement user behaviour analytics and separation of duties. NIST SP 800-161r1 AT-2(2) specifically requires awareness training to address insider threat risks within the supply chain context."),
            ],
        },
        "C": {
            "title": "Infrastructure Security",
            "nist_refs": "NIST SP 800-53r5: SC-7, SC-8, SI-2, SI-3, RA-5 / NIST SP 800-161r1: AC-17, SI-2",
            "recs": [
                ("Implement network segmentation & DMZ", "Separate web, application, and database tiers with a DMZ for internet-facing services. NIST SP 800-53r5 SC-7 requires boundary protection controls including managed interfaces and sub-networks (DMZ) for publicly accessible system components."),
                ("Enforce encrypted communications only", "Disable deprecated protocols (TLS 1.0/1.1, FTP, Telnet) and enforce TLS 1.2+ for all inter-system communication. Per NIST SP 800-53r5 SC-8, cryptographic mechanisms must protect the confidentiality and integrity of all transmitted information."),
                ("Mandate MFA for all remote access", "Enforce multi-factor authentication for VPN, remote desktop, and cloud console access. NIST SP 800-161r1 AC-17 specifies that remote access to supply chain systems must employ MFA, be limited to vetted personnel, and be contractually defined with restrictions by location and business hours."),
                ("Define SLA-based patch management", "Establish documented SLA timelines: Critical patches within 24h, High within 7 days, Medium within 30 days. NIST SP 800-53r5 SI-2 requires timely remediation of information system flaws; SP 800-161r1 SI-7 requires integrity verification of patches, including digital signature validation."),
                ("Deploy SIEM for centralized monitoring", "Aggregate logs in a SIEM and establish 24x7 SOC coverage. NIST SP 800-53r5 SI-4 requires continuous monitoring of information systems; NIST SP 800-161r1 AU-2 requires auditable supply chain events to be captured, correlated, and reviewed on an ongoing basis."),
            ],
        },
        "D": {
            "title": "Data Protection",
            "nist_refs": "NIST SP 800-53r5: SC-8, SC-12, SC-28, MP-6 / NIST SP 800-161r1: AC-4, PT-1",
            "recs": [
                ("Classify and label all data assets", "Implement a data classification policy aligned to sensitivity tiers. NIST SP 800-53r5 RA-2 requires information classification; SP 800-161r1 AC-4(19) requires validation of data metadata to ensure proper handling within the supply chain."),
                ("Encrypt data at rest, in transit & backup", "Apply AES-256 for data at rest and TLS 1.2+ in transit. NIST SP 800-53r5 SC-28 mandates protection of information at rest using cryptographic mechanisms; SC-8 covers data in transit. All backup media must also be encrypted prior to storage."),
                ("Implement formal key management", "Establish a documented key lifecycle (generation, rotation, revocation, destruction) with dual control. NIST SP 800-53r5 SC-12 requires cryptographic key establishment and management procedures; no single individual should have sole access to complete key material."),
                ("Enforce secure data destruction", "Define retention schedules and secure destruction procedures per NIST SP 800-53r5 MP-6, which requires sanitizing or destroying media containing information prior to disposal or reuse, including backup data at contract termination."),
                ("Deploy DLP controls", "Implement Data Loss Prevention to detect and prevent unauthorized exfiltration. NIST SP 800-161r1 AC-4 requires information flow controls across supply chain boundaries, including physical or logical separation to prevent unauthorized release of enterprise data."),
            ],
        },
        "E": {
            "title": "Access Management",
            "nist_refs": "NIST SP 800-53r5: AC-2, AC-3, AC-6, IA-2, IA-5 / NIST SP 800-161r1: AC-2, AC-6(6), IA-2",
            "recs": [
                ("Enforce need-to-know access provisioning", "Grant access only through a documented, approved request process. NIST SP 800-161r1 AC-2 requires unique contractor accounts with access that does not exceed the period of performance, and privileged accounts only for appropriately vetted personnel."),
                ("Conduct periodic access reviews", "Review all user access at least quarterly. NIST SP 800-53r5 AC-2(6) requires dynamic access management with regular review; SP 800-161r1 AC-3(8) requires prompt revocation when access is no longer needed, including credential retirement and account disabling."),
                ("Enforce organization-wide MFA", "Mandate MFA for all accounts, with stronger enforcement for privileged and remote users. NIST SP 800-53r5 IA-2(1) requires MFA for privileged access; IA-2(2) extends this to non-privileged accounts accessing sensitive systems."),
                ("Manage privileged access via PAM", "Implement a Privileged Access Management solution with JIT access, session recording, and dual control. NIST SP 800-161r1 AC-6(6) prohibits non-enterprise users from having privileged access and requires least-privilege mechanisms defining what is accessible, for what duration, and by whom."),
                ("Use strong credential storage", "Store all passwords using salted hashing (bcrypt, Argon2). NIST SP 800-53r5 IA-5 requires authenticator management practices that prevent plain-text storage; authenticators must be changed prior to delivery (IA-5(5)) and managed through federated credential controls where applicable."),
            ],
        },
        "F": {
            "title": "Application Security",
            "nist_refs": "NIST SP 800-53r5: SA-3, SA-8, SA-11, SA-15 / NIST SP 800-161r1: SA-3, SA-8, SI-7",
            "recs": [
                ("Embed security in the SDLC", "Integrate security requirements into every SDLC phase. NIST SP 800-53r5 SA-3 requires a system development life cycle with defined information security roles, and SA-15 requires a development process that addresses security during requirements, design, development, testing, and operations."),
                ("Mandate pre-launch penetration testing", "Conduct penetration testing and secure code review before any application goes live. NIST SP 800-53r5 SA-11 requires developer testing including security assessments; SP 800-161r1 SA-8 requires anticipating misuse scenarios in architecture and design."),
                ("Adopt secure coding standards", "Formally adopt OWASP Top 10 and SANS CWE Top 25. NIST SP 800-53r5 SA-8 requires security and privacy engineering principles to be applied, limiting privilege levels of critical elements and designing to reduce opportunities to exploit vulnerabilities."),
                ("Implement API security controls", "Enforce OAuth 2.0, rate limiting, and input validation on all APIs. NIST SP 800-53r5 SC-7 boundary protection principles apply to API endpoints; SA-8 requires controlling the number and privilege levels of interfaces exposed to external parties."),
                ("Maintain SBOM and SCA processes", "Generate and maintain a Software Bill of Materials and scan dependencies for vulnerabilities. Per NIST SP 800-161r1 SI-7, integrity of software components must be systematically tested and verified, including hash/signature validation of components from external repositories."),
            ],
        },
        "G": {
            "title": "System Security",
            "nist_refs": "NIST SP 800-53r5: AU-2, AU-3, AU-6, AU-12, SI-4, RA-5 / NIST SP 800-161r1: AU-2, AU-6, AU-16",
            "recs": [
                ("Generate comprehensive audit trails", "Enable audit logging for all admin actions, authentication events, and access changes. NIST SP 800-53r5 AU-2 requires event logging for user actions, failed logins, and account management; AU-3 specifies the required content of audit records including user ID, event type, date/time, and success/failure."),
                ("Protect logs against tampering", "Store logs in write-once or SIEM-forwarded storage with integrity monitoring. NIST SP 800-53r5 AU-9 protects audit information from unauthorized access, modification, and deletion; NIST SP 800-161r1 AU-10 requires non-repudiation techniques to protect the originality and integrity of audit records."),
                ("Review logs via automated SIEM correlation", "Implement SIEM with automated correlation rules for anomaly detection. NIST SP 800-161r1 AU-6 requires supply chain and information security audit events to be filtered, correlated, and reported; log review frequency must be adjusted based on vendor risk profile changes."),
                ("Conduct quarterly vulnerability scanning", "Run automated vulnerability scans at least quarterly with SLA-tracked remediation. NIST SP 800-53r5 RA-5 requires vulnerability scanning at defined frequencies; findings must be remediated within organization-defined time periods commensurate with their risk rating."),
                ("Retain logs for at least 12 months", "Maintain audit trail history for a minimum of 12 months with 3 months online. NIST SP 800-161r1 AU-16 requires cross-organizational audit logging with service-level agreements governing sharing and retention of audit information between the enterprise and its service providers."),
            ],
        },
        "H": {
            "title": "Email Security",
            "nist_refs": "NIST SP 800-53r5: SI-3, SI-8, SC-8 / NIST CSF 2.0: PR.PS",
            "recs": [
                ("Deploy email gateway with sandboxing", "Implement an email security gateway with URL rewriting and attachment sandboxing. NIST SP 800-53r5 SI-8 requires spam and malicious content protection for inbound and outbound email; SI-3 requires malicious code protection at entry and exit points."),
                ("Implement SPF, DKIM & DMARC", "Deploy email authentication protocols with at minimum a DMARC quarantine policy. NIST SP 800-53r5 SC-8 requires mechanisms protecting the integrity of transmitted information; email authentication prevents domain spoofing and impersonation attacks."),
                ("Encrypt sensitive email transmissions", "Encrypt outbound email attachments containing confidential or restricted data. Per NIST SP 800-53r5 SC-8(1), cryptographic mechanisms must be applied to protect the confidentiality of information during transmission over external networks."),
                ("Run annual phishing simulations", "Conduct simulated phishing exercises at least annually to measure awareness. NIST SP 800-53r5 AT-2(4) specifically requires training on suspicious communications and anomalous system behavior, with completion tracking and results used to update the training program."),
            ],
        },
        "I": {
            "title": "Mobile Devices",
            "nist_refs": "NIST SP 800-53r5: AC-19, MP-5, SC-28 / NIST SP 800-161r1: AC-19",
            "recs": [
                ("Enforce full-disk encryption on all laptops", "Mandate full-disk encryption (BitLocker, FileVault, or equivalent) enforced by policy. NIST SP 800-53r5 SC-28 requires protection of information at rest; MP-5 requires protection during transport, including on portable storage devices and laptops."),
                ("Enrol all devices in MDM", "Enrol all corporate and BYOD devices in a Mobile Device Management solution. NIST SP 800-53r5 AC-19 requires access control for mobile devices including configuration management, preventing unauthorized connection, and enforcing minimum security requirements before allowing corporate access."),
                ("Implement remote wipe capability", "Ensure all enrolled devices can be remotely locked and wiped via MDM. Per NIST SP 800-53r5 AC-19, organizations must be able to remotely wipe devices before returning to personally owned state; this is critical for data protection on lost or stolen devices."),
                ("Block non-compliant and rooted devices", "Automatically block jailbroken, rooted, or non-compliant devices from corporate resources. NIST SP 800-161r1 AC-19 requires access control for mobile devices accessing supply chain systems, with configuration management ensuring devices meet minimum security standards before connection."),
            ],
        },
        "J": {
            "title": "Incident Response",
            "nist_refs": "NIST SP 800-53r5: IR-2, IR-3, IR-4, IR-6, IR-8 / NIST SP 800-161r1: IR-4, IR-5, IR-6(3), IR-8",
            "recs": [
                ("Develop & maintain a documented IRP", "Establish a formal Incident Response Plan with defined roles and escalation paths. NIST SP 800-161r1 IR-8 requires the IRP to include information-sharing responsibilities with critical suppliers; the plan must cover supply chain-specific incidents and define coordination protocols with third parties."),
                ("Test the IRP at least annually", "Conduct tabletop exercises or simulations annually and document lessons learned. NIST SP 800-53r5 IR-3 requires testing of the incident response capability using defined exercises; results must feed back into plan updates and training improvements."),
                ("Define a contractual incident notification SLA", "Agree on a notification SLA (recommended: confirmed breaches within 4 hours). NIST SP 800-161r1 IR-6(3) requires incident reporting from suppliers to be protected in transmission and received only by approved individuals; reporting escalations must be clearly defined in the contract."),
                ("Establish 24x7 SOC coverage", "Maintain or contract 24x7 security operations for continuous incident detection. NIST SP 800-53r5 IR-4 requires an incident handling capability that includes preparation, detection, analysis, containment, eradication, and recovery; SP 800-161r1 IR-7(1) specifies that agreements must identify third-party incident response assistance conditions."),
                ("Include suppliers in incident response", "Integrate key suppliers into tabletop exercises and escalation matrices. NIST SP 800-161r1 IR-4(11) recommends that integrated incident response teams include forensics capability and, where practical, suppliers and external service providers with geographical representation."),
            ],
        },
        "K": {
            "title": "Cloud Services",
            "nist_refs": "NIST SP 800-53r5: SA-9, SC-7, SC-12, AU-2 / NIST SP 800-161r1: SA-9, AC-20",
            "recs": [
                ("Verify CSP third-party certifications", "Confirm CSP holds current ISO 27001, SOC 2 Type 2, or PCI-DSS certification. NIST SP 800-161r1 SA-9 requires periodic revalidation of external service provider adherence to security requirements via certifications or third-party assessments commensurate with criticality."),
                ("Confirm tenant data isolation", "Obtain and review the Shared Responsibility Matrix from the CSP. NIST SP 800-53r5 SA-9(5) requires that external service providers separate organizational information from that of other customers; multi-tenancy isolation must be documented and verified."),
                ("Enforce least privilege for cloud admin access", "Restrict hypervisor and admin console access to vetted personnel with MFA and least privilege. NIST SP 800-161r1 AC-20(1) limits authorized use of external systems; privileged access to CSP management planes must meet the same standards as internal privileged access controls."),
                ("Implement formal cloud key management", "Use CSP-native key management (AWS KMS, Azure Key Vault) with documented lifecycle. NIST SP 800-53r5 SC-12 requires cryptographic key establishment; key generation, distribution, storage, and destruction must follow a documented procedure with no single individual holding complete key material."),
                ("Require data deletion certificates", "Include contractual commitment for permanent deletion of all data at termination with written evidence. NIST SP 800-53r5 MP-6 requires media sanitization; for cloud services this must be contractually defined to ensure that all company data is purged from systems, storage, and backups upon termination."),
            ],
        },
        "L": {
            "title": "Business Continuity",
            "nist_refs": "NIST SP 800-53r5: CP-2, CP-4, CP-6, CP-9 / NIST SP 800-161r1: CP-2, CP-6, CP-7, CP-8",
            "recs": [
                ("Develop and approve a formal BCP", "Establish a formally approved Business Continuity Plan with defined RTO and RPO targets. NIST SP 800-53r5 CP-2 requires a contingency plan addressing essential missions, recovery objectives, and full system reconstitution; it must be reviewed at defined frequencies and approved by authorized officials."),
                ("Test the BCP at least annually", "Conduct annual BCP and DR tests (tabletop, failover, or full simulation). NIST SP 800-53r5 CP-4 requires testing of the contingency plan to validate effectiveness and identify gaps; test results must be documented and used to update the plan."),
                ("Validate RTO/RPO through actual testing", "Document RTO and RPO targets and verify them through actual failover testing. NIST SP 800-53r5 CP-2(1) requires coordination with external service providers; SP 800-161r1 CP-7 requires that alternative processing sites managed by service providers apply appropriate supply chain cybersecurity controls."),
                ("Implement off-site encrypted backups", "Store critical data backups securely off-site or in a separate cloud region. NIST SP 800-53r5 CP-6 requires an alternate storage site with appropriate controls; CP-9 requires information system backup including user-level data, system-level data, and system documentation, stored separately from the primary site."),
                ("Include suppliers in continuity planning", "Ensure critical suppliers are included in contingency plans and tested. NIST SP 800-161r1 CP-8(4) requires that telecommunications service provider contingency plans provide separation in infrastructure, service, process, and personnel to support supply chain resilience."),
            ],
        },
        "M": {
            "title": "Supply Chain & Physical Security",
            "nist_refs": "NIST SP 800-161r1: SR-1, SR-3, SR-6, PE-2 / NIST SP 800-53r5: SR-5, SR-6, PE-2, PE-3",
            "recs": [
                ("Establish a formal TPRM program", "Implement a documented Third-Party/Vendor Risk Management program. NIST SP 800-161r1 requires a C-SCRM governance structure with a dedicated PMO or equivalent function; vendor risk assessments must be performed at enterprise, mission, and operational levels (SP 800-161r1 Sec. 2)."),
                ("Inventory & disclose all sub-processors", "Identify and disclose all fourth parties with access to company data. NIST SP 800-161r1 requires supply chain visibility including identification of sub-tier suppliers; enterprises must flow C-SCRM control requirements down to prime contractors with requirements to further flow them to relevant sub-tier contractors."),
                ("Flow security requirements to sub-processors", "Include equivalent security obligations in all sub-processor contracts. NIST SP 800-161r1 SA-9 mandates satisfaction of applicable security requirements as a qualifying condition for award; contractual terms must address roles, responsibilities, and actions for responding to supply chain risk incidents."),
                ("Enforce physical access controls", "Restrict data centre access to authorized personnel with multi-factor physical controls. NIST SP 800-53r5 PE-2 requires physical access authorizations; PE-3 requires physical access control at all entry/exit points with controlled visitor access, escort requirements, and maintained access logs."),
                ("Assess sub-processors at least annually", "Conduct security assessments or evidence reviews for critical sub-processors annually. NIST SP 800-161r1 requires periodic revalidation of supplier adherence to security requirements; acceptable methods include certifications, site visits, third-party assessments, or self-attestation commensurate with criticality."),
            ],
        },
        "N": {
            "title": "AI & Emerging Technology Risk",
            "nist_refs": "NIST AI RMF 1.0: GOVERN, MAP, MEASURE / NIST CSF 2.0: GV.OC / NIST SP 800-53r5: SA-8, AU-2",
            "recs": [
                ("Establish a formal AI governance policy", "Publish an AI usage policy covering acceptable use, prohibited inputs, output handling, and accountability, approved by CISO/DPO. The NIST AI Risk Management Framework (AI RMF 1.0) GOVERN function requires organizations to establish policies, processes, and accountability structures for AI risk management before AI systems are deployed."),
                ("Conduct AI risk assessments", "Perform a formal AI risk assessment for all AI components in scope, reviewed annually. NIST AI RMF MAP function requires organizations to categorize AI risks by likelihood and impact; high-risk AI use cases (credit, fraud, identity) must undergo bias testing and explainability review per regulatory requirements."),
                ("Require DPA from AI service providers", "Obtain Data Processing Agreements from all third-party AI providers confirming no model training on client data. NIST SP 800-161r1 PT-1 requires contracts to specify what data will be shared, which personnel may access it, applicable controls, retention periods, and data handling at contract end."),
                ("Implement PII controls for AI inputs", "Deploy masking or tokenization to prevent PII from entering AI models unnecessarily. NIST AI RMF MEASURE function requires measurement of data quality, relevance, and bias; technical controls must prevent sensitive data from being exposed to AI systems where not required for the intended use case."),
                ("Protect AI systems from adversarial attacks", "Deploy prompt injection detection and output guardrails; enforce RBAC and MFA for all AI system access. NIST SP 800-53r5 SA-8 requires security engineering principles to anticipate maximum possible misuse scenarios; adversarial inputs and prompt injection are recognized attack vectors that must be addressed in AI system design."),
                ("Maintain tamper-protected AI audit logs", "Log all AI interactions (inputs, outputs, user identity, timestamps) in tamper-protected storage. NIST SP 800-53r5 AU-2 and AU-12 require event logging and audit record generation for system interactions; for AI systems this extends to model inputs and outputs to support forensic investigation of misuse or data leakage incidents."),
            ],
        },
    }

    buf = io.BytesIO()
    PAGE_W, PAGE_H = A4
    MARGIN = 16 * mm
    usable_w = PAGE_W - 2 * MARGIN

    if not assessment_date:
        assessment_date = datetime.date.today().strftime("%d %B %Y")
    vendor_line = vendor_name if vendor_name else "—"

    # ── colour palette ─────────────────────────────────────────────────────────
    C_RED    = colors.HexColor("#A32D2D")
    C_AMB    = colors.HexColor("#854F0B")
    C_BLU    = colors.HexColor("#185FA5")
    C_GRY    = colors.HexColor("#6c757d")
    C_BDR    = colors.HexColor("#e9ecef")
    C_HDR_BG = colors.HexColor("#1a1a2e")
    C_TBL_BG = colors.HexColor("#f8f9fa")
    C_REC_BG = colors.HexColor("#F0F4FF")
    C_REC_BD = colors.HexColor("#c7d2fe")
    TIER_FG  = {"Critical": C_RED, "High": C_AMB, "Medium": C_BLU}

    def _s(name, **kw):
        s = ParagraphStyle(name)
        for k, v in kw.items():
            setattr(s, k, v)
        return s

    s_title   = _s("title",   fontSize=13, fontName="Helvetica-Bold",  textColor=colors.white,            alignment=TA_LEFT,   leading=17)
    s_body    = _s("body",    fontSize=8,  fontName="Helvetica",        textColor=colors.HexColor("#333"), leading=11)
    s_small   = _s("small",   fontSize=7,  fontName="Helvetica",        textColor=C_GRY,                   leading=10)
    s_ref     = _s("ref",     fontSize=7.5,fontName="Helvetica-Bold",   textColor=C_GRY)
    s_sec     = _s("sec",     fontSize=7.5,fontName="Helvetica-Bold",   textColor=C_GRY,                   spaceBefore=4,  spaceAfter=2)
    s_note    = _s("note",    fontSize=7.5,fontName="Helvetica-Oblique",textColor=C_GRY,                   leading=11)
    s_footer  = _s("footer",  fontSize=7,  fontName="Helvetica",        textColor=C_GRY,                   alignment=TA_CENTER)
    s_nist    = _s("nist",    fontSize=6.5,fontName="Helvetica-Oblique",textColor=C_BLU,                   leading=9)
    s_rec_hd  = _s("rec_hd",  fontSize=7.5,fontName="Helvetica-Bold",  textColor=colors.HexColor("#1a1a2e"), leading=10)
    s_rec_bod = _s("rec_bod", fontSize=7,  fontName="Helvetica",        textColor=colors.HexColor("#444"), leading=10)
    s_dom_hd  = _s("dom_hd",  fontSize=9,  fontName="Helvetica-Bold",  textColor=colors.white,             leading=12)

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN,  bottomMargin=MARGIN,
    )
    story = []

    # ── shared helpers ─────────────────────────────────────────────────────────
    def _footer_row(page_label):
        ft = Table(
            [[
                Paragraph(f"TPCRA v3.0  |  For vendor assessment use only  |  Internal — Confidential  |  {page_label}", s_footer),
                Paragraph(f"Generated: {assessment_date}", _s("fr", fontSize=7, fontName="Helvetica", textColor=C_GRY, alignment=TA_RIGHT)),
            ]],
            colWidths=[usable_w * 0.72, usable_w * 0.28],
        )
        ft.setStyle(TableStyle([
            ("TOPPADDING",  (0,0),(-1,-1), 5),
            ("LINEABOVE",   (0,0),(-1,-1), 0.5, C_BDR),
        ]))
        return ft

    def _header_banner(subtitle=""):
        title_text = "Third-Party Cyber Risk Assessment  —  Gap Summary Report"
        if subtitle:
            title_text += f"  |  {subtitle}"
        t = Table([[Paragraph(title_text, s_title)]], colWidths=[usable_w])
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), C_HDR_BG),
            ("TOPPADDING",   (0,0),(-1,-1), 10),
            ("BOTTOMPADDING",(0,0),(-1,-1), 10),
            ("LEFTPADDING",  (0,0),(-1,-1), 12),
            ("RIGHTPADDING", (0,0),(-1,-1), 12),
        ]))
        return t

    def _meta_row():
        t = Table(
            [[
                Paragraph(f"<b>Vendor:</b> {vendor_line}", s_body),
                Paragraph(f"<b>Assessment date:</b> {assessment_date}", s_body),
                Paragraph("<b>Framework:</b> TPCRA v3.0", s_body),
                Paragraph("<b>Classification:</b> Confidential", s_body),
            ]],
            colWidths=[usable_w*0.30, usable_w*0.26, usable_w*0.22, usable_w*0.22],
        )
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), C_TBL_BG),
            ("BOX",          (0,0),(-1,-1), 0.5, C_BDR),
            ("TOPPADDING",   (0,0),(-1,-1), 6),
            ("BOTTOMPADDING",(0,0),(-1,-1), 6),
            ("LEFTPADDING",  (0,0),(-1,-1), 8),
            ("RIGHTPADDING", (0,0),(-1,-1), 8),
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
        ]))
        return t

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 1 — Gap overview
    # ══════════════════════════════════════════════════════════════════════════
    story.append(_header_banner("Page 1 of 2 — Gap Overview"))
    story.append(Spacer(1, 2.5*mm))
    story.append(_meta_row())
    story.append(Spacer(1, 3*mm))

    # KPI strip
    total_c  = sum(d["critical"] for d in gap_db)
    total_h  = sum(d["high"]     for d in gap_db)
    total_m  = sum(d["medium"]   for d in gap_db)
    flagged  = sum(1 for d in gap_db if d["risk"] in ("Critical","High"))
    n_dom    = len(gap_db)

    def _kpi(val, label, color="#333333"):
        return Paragraph(
            f'<font size="13" color="{color}"><b>{val}</b></font><br/>'
            f'<font size="7" color="#6c757d">{label}</font>',
            _s(f"kpi{val}{label[:3]}", fontName="Helvetica", alignment=TA_CENTER, leading=17),
        )

    kpi_tbl = Table(
        [[
            _kpi(n_dom,   "Domains assessed"),
            _kpi(total_c, "Critical controls", "#A32D2D"),
            _kpi(total_h, "High controls",     "#854F0B"),
            _kpi(total_m, "Medium controls",   "#185FA5"),
            _kpi(flagged, "Domains flagged",   "#A32D2D"),
        ]],
        colWidths=[usable_w/5]*5,
    )
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), colors.white),
        ("BOX",          (0,0),(-1,-1), 0.5, C_BDR),
        ("INNERGRID",    (0,0),(-1,-1), 0.5, C_BDR),
        ("TOPPADDING",   (0,0),(-1,-1), 8),
        ("BOTTOMPADDING",(0,0),(-1,-1), 8),
        ("ALIGN",        (0,0),(-1,-1), "CENTER"),
        ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 3.5*mm))

    # Domain gap table
    story.append(Paragraph("SECURITY DOMAIN GAP OVERVIEW", s_sec))
    col_w = [usable_w*0.04, usable_w*0.20, usable_w*0.07, usable_w*0.06, usable_w*0.09, usable_w*0.54]
    tbl_data = [[
        Paragraph("#",              s_small),
        Paragraph("Domain",         s_small),
        Paragraph("Critical",       s_small),
        Paragraph("High",           s_small),
        Paragraph("Rating",         s_small),
        Paragraph("Priority gaps (top 2 critical/high items)", s_small),
    ]]
    for i, d in enumerate(gap_db):
        top_gaps = [g for g in d["gaps"] if g["tier"] in ("Critical","High")][:2]
        gap_lines = "\n".join(
            f'• {g["ref"]}: {g["text"][:80]}{"..." if len(g["text"])>80 else ""}'
            for g in top_gaps
        ) if top_gaps else "No critical/high gaps identified."
        risk_fg = TIER_FG.get(d["risk"], C_GRY)
        tbl_data.append([
            Paragraph(d["id"],            s_ref),
            Paragraph(d["name"],          s_body),
            Paragraph(str(d["critical"]), _s(f"rc{i}", fontName="Helvetica-Bold", fontSize=8, textColor=C_RED, alignment=TA_CENTER)),
            Paragraph(str(d["high"]),     _s(f"rh{i}", fontName="Helvetica-Bold", fontSize=8, textColor=C_AMB, alignment=TA_CENTER)),
            Paragraph(d["risk"],          _s(f"rr{i}", fontName="Helvetica-Bold", fontSize=7, textColor=risk_fg, alignment=TA_CENTER)),
            Paragraph(gap_lines,          _s(f"rg{i}", fontName="Helvetica", fontSize=7, textColor=colors.HexColor("#333"), leading=10)),
        ])

    dom_tbl = Table(tbl_data, colWidths=col_w, repeatRows=1)
    ts = [
        ("BACKGROUND",   (0,0),(-1,0),  C_TBL_BG),
        ("BOX",          (0,0),(-1,-1), 0.5, C_BDR),
        ("INNERGRID",    (0,0),(-1,-1), 0.3, C_BDR),
        ("TOPPADDING",   (0,0),(-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ("LEFTPADDING",  (0,0),(-1,-1), 5),
        ("RIGHTPADDING", (0,0),(-1,-1), 5),
        ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ("ALIGN",        (2,1),(4,-1),  "CENTER"),
    ]
    for r in range(1, len(tbl_data)):
        if r % 2 == 0:
            ts.append(("BACKGROUND", (0,r),(-1,r), colors.HexColor("#fafafa")))
    dom_tbl.setStyle(TableStyle(ts))
    story.append(dom_tbl)
    story.append(Spacer(1, 3*mm))

    # Expectations note
    note_tbl = Table(
        [[Paragraph(
            "<b>Vendor expectations:</b> All Critical-tier control gaps must be remediated or a documented "
            "compensating control provided before onboarding can proceed. High-tier gaps require a remediation "
            "plan with committed timelines submitted within 30 days. Evidence must be provided per the TPCRA v3.0 "
            "Evidence Checklist. All information provided is treated as confidential and used solely for "
            "third-party risk assessment purposes.",
            s_note,
        )]],
        colWidths=[usable_w],
    )
    note_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), C_REC_BG),
        ("BOX",          (0,0),(-1,-1), 0.5, C_REC_BD),
        ("TOPPADDING",   (0,0),(-1,-1), 7),
        ("BOTTOMPADDING",(0,0),(-1,-1), 7),
        ("LEFTPADDING",  (0,0),(-1,-1), 9),
        ("RIGHTPADDING", (0,0),(-1,-1), 9),
    ]))
    story.append(note_tbl)
    story.append(Spacer(1, 3*mm))
    story.append(_footer_row("Page 1 of 2"))

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2 — NIST-grounded recommendations per domain
    # ══════════════════════════════════════════════════════════════════════════
    from reportlab.platypus import PageBreak
    story.append(PageBreak())

    story.append(_header_banner("Page 2 of 2 — Recommendations (NIST-grounded)"))
    story.append(Spacer(1, 2.5*mm))
    story.append(_meta_row())
    story.append(Spacer(1, 3*mm))

    intro_tbl = Table(
        [[Paragraph(
            "<b>How to use this page:</b> Each domain below lists actionable recommendations with explicit "
            "references to NIST SP 800-53r5, NIST SP 800-161r1 (C-SCRM), NIST CSF 2.0, and/or NIST AI RMF 1.0. "
            "Vendors should address all Critical-rated domains as a priority. Each recommendation maps to "
            "specific control families to guide implementation and evidence collection.",
            s_note,
        )]],
        colWidths=[usable_w],
    )
    intro_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), C_REC_BG),
        ("BOX",          (0,0),(-1,-1), 0.5, C_REC_BD),
        ("TOPPADDING",   (0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",  (0,0),(-1,-1), 9),
        ("RIGHTPADDING", (0,0),(-1,-1), 9),
    ]))
    story.append(intro_tbl)
    story.append(Spacer(1, 3*mm))

    # Two-column layout for domain recommendations
    HALF = (usable_w - 3*mm) / 2

    def _domain_rec_cell(d, rec_data):
        risk_fg = TIER_FG.get(d["risk"], C_GRY)
        # Domain header
        hdr = Table(
            [[
                Paragraph(f"{d['id']} — {rec_data['title']}", s_dom_hd),
                Paragraph(d["risk"], _s(f"dh{d['id']}", fontSize=7.5, fontName="Helvetica-Bold",
                                        textColor=colors.white, alignment=TA_RIGHT, leading=12)),
            ]],
            colWidths=[HALF*0.78, HALF*0.22],
        )
        hdr_bg = C_RED if d["risk"] == "Critical" else (C_AMB if d["risk"] == "High" else C_BLU)
        hdr.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), hdr_bg),
            ("TOPPADDING",   (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",  (0,0),(-1,-1), 6),
            ("RIGHTPADDING", (0,0),(-1,-1), 6),
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
        ]))

        # NIST ref tag
        nist_tag = Table(
            [[Paragraph(rec_data["nist_refs"], s_nist)]],
            colWidths=[HALF],
        )
        nist_tag.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), colors.HexColor("#EEF2FF")),
            ("TOPPADDING",   (0,0),(-1,-1), 3),
            ("BOTTOMPADDING",(0,0),(-1,-1), 3),
            ("LEFTPADDING",  (0,0),(-1,-1), 6),
            ("RIGHTPADDING", (0,0),(-1,-1), 6),
        ]))

        # Recommendation rows
        rec_rows = []
        for j, (heading, detail) in enumerate(rec_data["recs"]):
            num_para = Paragraph(str(j+1), _s(f"rn{d['id']}{j}", fontSize=7.5, fontName="Helvetica-Bold",
                                               textColor=colors.white, alignment=TA_CENTER))
            num_cell = Table([[num_para]], colWidths=[5*mm])
            num_cell.setStyle(TableStyle([
                ("BACKGROUND",   (0,0),(-1,-1), hdr_bg),
                ("TOPPADDING",   (0,0),(-1,-1), 2),
                ("BOTTOMPADDING",(0,0),(-1,-1), 2),
                ("LEFTPADDING",  (0,0),(-1,-1), 0),
                ("RIGHTPADDING", (0,0),(-1,-1), 0),
                ("VALIGN",       (0,0),(-1,-1), "TOP"),
            ]))
            text_cell = Table(
                [[Paragraph(heading, s_rec_hd)],
                 [Paragraph(detail,  s_rec_bod)]],
                colWidths=[HALF - 5*mm - 8],
            )
            text_cell.setStyle(TableStyle([
                ("TOPPADDING",   (0,0),(-1,-1), 1),
                ("BOTTOMPADDING",(0,0),(-1,-1), 1),
                ("LEFTPADDING",  (0,0),(-1,-1), 4),
                ("RIGHTPADDING", (0,0),(-1,-1), 2),
            ]))
            row_tbl = Table([[num_cell, text_cell]], colWidths=[5*mm, HALF - 5*mm])
            row_tbl.setStyle(TableStyle([
                ("TOPPADDING",   (0,0),(-1,-1), 2),
                ("BOTTOMPADDING",(0,0),(-1,-1), 2),
                ("LEFTPADDING",  (0,0),(-1,-1), 3),
                ("RIGHTPADDING", (0,0),(-1,-1), 3),
                ("VALIGN",       (0,0),(-1,-1), "TOP"),
                ("LINEBELOW",    (0,0),(-1,-1), 0.3, C_BDR),
            ]))
            rec_rows.append(row_tbl)

        # Wrap all into a card
        card_content = [hdr, nist_tag] + rec_rows
        card = Table([[item] for item in card_content], colWidths=[HALF])
        card.setStyle(TableStyle([
            ("BOX",          (0,0),(-1,-1), 0.5, C_BDR),
            ("TOPPADDING",   (0,0),(-1,-1), 0),
            ("BOTTOMPADDING",(0,0),(-1,-1), 0),
            ("LEFTPADDING",  (0,0),(-1,-1), 0),
            ("RIGHTPADDING", (0,0),(-1,-1), 0),
        ]))
        return card

    # Lay out domains in 2-column pairs
    domains_with_recs = [(d, NIST_RECS[d["id"]]) for d in gap_db if d["id"] in NIST_RECS]
    for i in range(0, len(domains_with_recs), 2):
        left_d,  left_r  = domains_with_recs[i]
        if i+1 < len(domains_with_recs):
            right_d, right_r = domains_with_recs[i+1]
            row = Table(
                [[_domain_rec_cell(left_d, left_r), _domain_rec_cell(right_d, right_r)]],
                colWidths=[HALF, HALF],
                hAlign="LEFT",
            )
        else:
            # Odd domain — full width spanning both columns
            row = Table(
                [[_domain_rec_cell(left_d, left_r), ""]],
                colWidths=[HALF, HALF],
                hAlign="LEFT",
            )
        row.setStyle(TableStyle([
            ("TOPPADDING",   (0,0),(-1,-1), 2),
            ("BOTTOMPADDING",(0,0),(-1,-1), 2),
            ("LEFTPADDING",  (0,0),(-1,-1), 0),
            ("RIGHTPADDING", (0,0),(-1,-1), 0),
            ("VALIGN",       (0,0),(-1,-1), "TOP"),
            ("COLPADDING",   (0,0),(-1,-1), 1.5*mm),
        ]))
        story.append(row)
        story.append(Spacer(1, 1.5*mm))

    story.append(Spacer(1, 2*mm))
    story.append(_footer_row("Page 2 of 2"))

    doc.build(story)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.header("Upload questionnaire")
    uploaded = st.file_uploader(
        "Use TPCRA v3.0 questionnaire",
        type=["xlsx", "xls"],
        help="Upload a completed TPCRA v3.0 questionnaire."
    )
    st.divider()
    st.download_button(
        "⬇ Download sample template",
        data=make_sample_excel(),
        file_name="TPCRA_v3.0_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download a minimal TPCRA v3.0 questionnaire template to fill in.",
        use_container_width=True,
    )

# ── Empty state ────────────────────────────────────────────────────────────────
if not uploaded:
    st.title("🔐 TPCRA v3.0 Risk Assessment Dashboard")
    st.markdown("Upload a completed TPCRA v3.0 questionnaire from the sidebar to generate the dashboard.")
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.info("**Overview**\nCompliance score, risk rating, response distribution, and domain scores across all 14 domains.")
    c2.info("**By domain**\nDrill into any of the A–N domains with per-question response cards, tier badges, and vendor remarks.")
    c3.info("**Gap summary**\nAll No / N/A / Partial / unanswered controls grouped by tier, with NIST-grounded recommendations and PDF export.")
    c4, c5, c6 = st.columns(3)
    c4.info("**Gap analysis**\nFlat filterable list of every gap item — by tier, response type, and domain — with CSV/Excel export.")
    c5.info("**Evidence checklist**\nEvidence submission status against the 14 required evidence items from the TPCRA v3.0 checklist.")
    c6.info("**Engagement info**\nVendor contact details, engagement description, data handling, and transmission methods from Part 1.")
    st.stop()

# ── Load workbook ──────────────────────────────────────────────────────────────
try:
    wb = load_workbook(uploaded, read_only=True, data_only=True)
except Exception as e:
    st.error(f"Could not open file: {e}")
    st.stop()

p1_data  = parse_part1(wb)
p2_data  = parse_part2(wb)
evidence = parse_evidence(wb)

if not p2_data or not p2_data.get("items"):
    st.error("No Part 2 question data found. Ensure the file has a 'Part 2' sheet matching the TPCRA v3.0 format.")
    st.stop()

contact  = extract_contact(p1_data.get("items", []))
p2_items = p2_data["items"]
domains  = p2_data["domains"]

# ── Header ─────────────────────────────────────────────────────────────────────
h1, h2 = st.columns([4, 1])
with h1:
    vendor_label = contact["vendor"] or "Vendor"
    st.title(f"🔐 {vendor_label}")
    if contact["rep"] or contact["email"]:
        st.caption(f"{contact['rep']}  ·  {contact['email']}")
    if contact["engagement"]:
        st.caption(f"Engagement: {contact['engagement']}")
with h2:
    st.caption(p2_data.get("title", "TPCRA v3.0"))

st.divider()

# ── KPI metrics ────────────────────────────────────────────────────────────────
answered = [i for i in p2_items if i["norm"] in ("Yes", "No", "Partial", "N/A")]
unanswered = [i for i in p2_items if i["norm"] == "—"]
n_yes  = sum(1 for i in p2_items if i["norm"] == "Yes")
n_no   = sum(1 for i in p2_items if i["norm"] == "No")
n_part = sum(1 for i in p2_items if i["norm"] == "Partial")
n_na   = sum(1 for i in p2_items if i["norm"] == "N/A")
n_unans = len(unanswered)
score  = compliance_score(p2_items)
rating_label, rating_color = risk_rating(score)

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
k1.metric("Total questions", len(p2_items))
k2.metric("✅ Yes",           n_yes)
k3.metric("❌ No",            n_no)
k4.metric("⚠️ Partial",       n_part)
k5.metric("➖ N/A",           n_na)
k6.metric("⬜ Unanswered",    n_unans)
k7.metric("Compliance score",  f"{score}%",
    delta=rating_label,
    delta_color="normal" if score >= 70 else ("off" if score >= 50 else "inverse"),
)
st.divider()

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_overview, tab_domain, tab_evidence, tab_part1, tab_gap_summary = st.tabs([
    "Overview", "By domain", "Evidence checklist", "Engagement info", "Gap summary"
])

# ══════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════
with tab_overview:
    col_bar, col_right = st.columns([3, 2])

    with col_bar:
        st.subheader("Compliance by domain")
        dom_rows = []
        for letter, dom in domains.items():
            items = dom["items"]
            if not items:
                continue
            sc = compliance_score(items)
            rl, rc = risk_rating(sc)
            dom_rows.append({
                "Domain":  dom["name"],
                "Yes":     sum(1 for i in items if i["norm"] == "Yes"),
                "Partial": sum(1 for i in items if i["norm"] == "Partial"),
                "No":      sum(1 for i in items if i["norm"] == "No"),
                "N/A":     sum(1 for i in items if i["norm"] == "N/A"),
                "Score":   sc,
                "Rating":  rl,
            })
        dom_df = pd.DataFrame(dom_rows)

        if not dom_df.empty:
            melt = dom_df.melt(
                id_vars="Domain", value_vars=["Yes", "Partial", "No", "N/A"],
                var_name="Response", value_name="Count"
            ).query("Count > 0")
            fig_bar = px.bar(
                melt, x="Count", y="Domain", color="Response",
                orientation="h",
                color_discrete_map=RESP_COLORS,
                category_orders={"Response": ["Yes", "Partial", "No", "N/A"]},
                labels={"Count": "Questions", "Domain": ""},
                height=max(380, len(dom_df) * 42),
            )
            fig_bar.update_layout(
                margin=dict(l=0, r=0, t=10, b=0),
                legend_title_text="Response",
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                yaxis={"categoryorder": "total ascending"},
                font=dict(size=12),
            )
            fig_bar.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.07)")
            st.plotly_chart(fig_bar, use_container_width=True)

    with col_right:
        st.subheader("Distribution")
        labels = [k for k, v in [("Yes",n_yes),("Partial",n_part),("No",n_no),("N/A",n_na)] if v > 0]
        values = [v for k, v in [("Yes",n_yes),("Partial",n_part),("No",n_no),("N/A",n_na)] if v > 0]
        fig_donut = go.Figure(go.Pie(
            labels=labels, values=values, hole=0.62,
            marker_colors=[RESP_COLORS[l] for l in labels],
            textinfo="label+percent",
            hovertemplate="%{label}: %{value}<extra></extra>",
        ))
        fig_donut.update_layout(
            margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
            height=260, paper_bgcolor="rgba(0,0,0,0)", font=dict(size=12),
        )
        st.plotly_chart(fig_donut, use_container_width=True)

        st.subheader("Domain scores")
        if not dom_df.empty:
            for _, row in dom_df.sort_values("Score").iterrows():
                _, rc = risk_rating(row["Score"])
                st.markdown(
                    f'<div style="margin-bottom:7px">'
                    f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px">'
                    f'<span style="color:#555">{row["Domain"]}</span>'
                    f'<span style="font-weight:600;color:{rc}">{row["Score"]}%</span></div>'
                    f'<div style="background:#f0f0f0;border-radius:4px;height:5px;overflow:hidden">'
                    f'<div style="width:{row["Score"]}%;height:100%;background:{rc};border-radius:4px"></div>'
                    f'</div></div>', unsafe_allow_html=True,
                )

# ══════════════════════════
# TAB 2 — BY DOMAIN
# ══════════════════════════
with tab_domain:
    domain_choices = [f"{l} — {domains[l]['name']}" for l in domains if domains[l]["items"]]
    if not domain_choices:
        st.info("No domain data found.")
    else:
        chosen = st.selectbox("Select domain", domain_choices)
        chosen_letter = chosen.split(" — ")[0]
        dom = domains[chosen_letter]
        items = dom["items"]

        sc = compliance_score(items)
        rl, rc = risk_rating(sc)
        n_y = sum(1 for i in items if i["norm"] == "Yes")
        n_n = sum(1 for i in items if i["norm"] == "No")
        n_p = sum(1 for i in items if i["norm"] == "Partial")
        n_a = sum(1 for i in items if i["norm"] == "N/A")

        dm1, dm2, dm3, dm4, dm5, dm6 = st.columns(6)
        dm1.metric("Questions", len(items))
        dm2.metric("✅ Yes",    n_y)
        dm3.metric("❌ No",     n_n)
        dm4.metric("⚠️ Partial", n_p)
        dm5.metric("➖ N/A",    n_a)
        dm6.metric("Score", f"{sc}%", delta=rl,
            delta_color="normal" if sc >= 70 else ("off" if sc >= 50 else "inverse"))

        st.divider()

        # Filter
        resp_f = st.multiselect(
            "Filter by response", ["Yes","No","Partial","N/A","—"],
            default=["Yes","No","Partial","N/A","—"], key="dom_resp_filter"
        )
        shown = [i for i in items if i["norm"] in resp_f]
        st.caption(f"Showing {len(shown)} of {len(items)} questions")

        current_sub = ""
        for item in shown:
            # Sub-section label
            if item["sub"] and item["sub"] != current_sub:
                current_sub = item["sub"]
                st.markdown(
                    f'<div style="margin:16px 0 6px;font-size:11px;font-weight:700;'
                    f'text-transform:uppercase;letter-spacing:0.06em;color:#6c757d">'
                    f'{current_sub}</div>', unsafe_allow_html=True
                )

            norm = item["norm"]
            bg = {"No": "#fff5f5", "Partial": "#fffaf5"}.get(norm, "transparent")
            border = {"No": "#f7c1c1", "Partial": "#FAC775"}.get(norm, "#e9ecef")
            tier_badge = tier_pill(item["tier"]) if item["tier"] else ""
            resp_badge = resp_pill(norm)
            key_line = item["key"]
            tier_part = "&nbsp;&nbsp;" + tier_badge if tier_badge else ""
            remarks_html = (
                '<div style="font-size:12px;color:#555;margin-top:6px;'
                'padding-top:6px;border-top:0.5px solid #e9ecef">'
                + item["other"] + "</div>"
            ) if item["other"] else ""
            card_html = (
                '<div style="padding:10px 14px;border-radius:8px;margin-bottom:6px;'
                'background:' + bg + ';border:0.5px solid ' + border + '">'
                '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px">'
                '<div style="flex:1">'
                '<div style="font-size:11px;color:#aaa;margin-bottom:3px">'
                + key_line + tier_part +
                '</div>'
                '<div style="font-size:13px;color:#333;line-height:1.55">' + item["question"] + '</div>'
                + remarks_html +
                '</div>'
                '<div style="flex-shrink:0;padding-top:2px">' + resp_badge + '</div>'
                '</div>'
                '</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)

            # Long free-text responses
            if item["response"] and len(item["response"]) > 50 and norm not in ("Yes","No","Partial","N/A","—"):
                with st.expander(f"View full response — {item['key']}"):
                    st.markdown(
                        f'<div style="font-size:13px;color:#444;line-height:1.6;white-space:pre-wrap">'
                        f'{item["response"]}</div>', unsafe_allow_html=True
                    )



# ══════════════════════════
# TAB 3 — EVIDENCE
# ══════════════════════════
with tab_evidence:
    if not evidence:
        st.info("No Evidence sheet found in the uploaded file.")
    else:
        submitted = sum(1 for e in evidence if e["status"] == "Submitted")
        pending   = len(evidence) - submitted

        ev1, ev2, ev3 = st.columns(3)
        ev1.metric("Total evidence items", len(evidence))
        ev2.metric("✅ Submitted",          submitted)
        ev3.metric("⏳ Pending",            pending)
        st.divider()

        rows_html = ""
        for ev in evidence:
            s = ev["status"]
            s_style = "background:#EAF3DE;color:#27500A" if s == "Submitted" else "background:#FAEEDA;color:#633806"
            rows_html += (
                "<tr>"
                '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#aaa;width:4%;vertical-align:top">' + str(ev['num']) + "</td>"
                '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;color:#333;width:26%;vertical-align:top;font-weight:500">' + str(ev['evidence']) + "</td>"
                '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#555;width:35%;vertical-align:top;line-height:1.5">' + str(ev['guidance']) + "</td>"
                '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;width:20%;vertical-align:top">' + str(ev['required_for']) + "</td>"
                '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;width:15%;vertical-align:top">'
                '<span class="response-pill" style="' + s_style + '">' + s + "</span></td>"
                "</tr>"
            )

        TH = "padding:9px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef"
        ev_table = (
            '<div style="border:1px solid #e9ecef;border-radius:10px;overflow:hidden">'
            '<table style="width:100%;border-collapse:collapse;table-layout:fixed">'
            '<thead><tr style="background:#f8f9fa">'
            '<th style="' + TH + ';width:4%">#</th>'
            '<th style="' + TH + ';width:26%">Evidence required</th>'
            '<th style="' + TH + ';width:35%">Guidance</th>'
            '<th style="' + TH + ';width:20%">Required for</th>'
            '<th style="' + TH + ';width:15%">Status</th>'
            "</tr></thead>"
            "<tbody>" + rows_html + "</tbody>"
            "</table></div>"
        )
        st.markdown(ev_table, unsafe_allow_html=True)


# ══════════════════════════
# TAB 4 — ENGAGEMENT INFO
# ══════════════════════════
with tab_part1:
    if not p1_data or not p1_data.get("items"):
        st.info("No Part 1 data found in the uploaded file. Ensure the file contains a 'Part 1' sheet.")
    else:
        p1_sections = p1_data.get("sections", {})
        if not p1_sections:
            sections_to_show = {"All questions": p1_data["items"]}
        else:
            sections_to_show = p1_sections

        HIDE_TIER_SECTIONS = {"SECTION 1 — CONTACT PERSON", "SECTION 2 — ENGAGEMENT INFORMATION"}

        for sec_name, sec_items in sections_to_show.items():
            if not sec_items:
                continue
            st.subheader(sec_name.replace("SECTION ", "").replace(" — ", " — ").title()
                         if "SECTION" in sec_name else sec_name)

            hide_tier = sec_name in HIDE_TIER_SECTIONS

            rows_html = ""
            for item in sec_items:
                r = item["response"] or "—"
                resp_preview = r if len(r) <= 100 else r[:97] + "…"
                tier_cell = (
                    ""
                    if hide_tier
                    else '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;width:10%;vertical-align:top">'
                         + (tier_pill(item["tier"]) if item.get("tier") else "")
                         + "</td>"
                )
                q_width  = "55%" if hide_tier else "48%"
                resp_width = "38%" if hide_tier else "35%"
                rows_html += (
                    "<tr>"
                    '<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#aaa;width:7%;vertical-align:top">' + item['key'] + "</td>"
                    f'<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;color:#333;width:{q_width};vertical-align:top;line-height:1.5">' + item['question'] + "</td>"
                    f'<td style="padding:9px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;color:#555;width:{resp_width};vertical-align:top;line-height:1.5">' + resp_preview + "</td>"
                    + tier_cell +
                    "</tr>"
                )

            TH2 = "padding:9px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef"
            tier_header = "" if hide_tier else f'<th style="{TH2};width:10%">Tier</th>'
            q_width_h   = "55%" if hide_tier else "48%"
            resp_width_h = "38%" if hide_tier else "35%"
            p1_table = (
                '<div style="border:1px solid #e9ecef;border-radius:10px;overflow:hidden;margin-bottom:1.5rem">'
                '<table style="width:100%;border-collapse:collapse;table-layout:fixed">'
                '<thead><tr style="background:#f8f9fa">'
                f'<th style="{TH2};width:7%">#</th>'
                f'<th style="{TH2};width:{q_width_h}">Question</th>'
                f'<th style="{TH2};width:{resp_width_h}">Response</th>'
                + tier_header +
                "</tr></thead>"
                "<tbody>" + rows_html + "</tbody>"
                "</table></div>"
            )
            st.markdown(p1_table, unsafe_allow_html=True)


# ══════════════════════════
# TAB 5 — GAP SUMMARY
# ══════════════════════════
with tab_gap_summary:

    # ── Response style helpers (reuse existing pill constants) ──────────────────
    GS_RISK_BADGE = {
        "Critical": "background:#FCEBEB;color:#791F1F;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
        "High":     "background:#FAEEDA;color:#633806;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
        "Medium":   "background:#E6F1FB;color:#0C447C;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
        "Low":      "background:#EAF3DE;color:#27500A;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
    }
    GS_RESP_BADGE = {
        "No":      "background:#FCEBEB;color:#A32D2D;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
        "N/A":     "background:#F1EFE8;color:#5F5E5A;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
        "Partial": "background:#FAEEDA;color:#633806;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
        "—":       "background:#F1EFE8;color:#888780;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600",
    }
    TIER_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3, "": 4}
    TIER_BORDER = {
        "Critical": "#f7c1c1",
        "High":     "#FAC775",
        "Medium":   "#B5D4F4",
        "Low":      "#C0DD97",
        "":         "#e9ecef",
    }
    TIER_BG = {
        "Critical": "#fff5f5",
        "High":     "#fffaf5",
        "Medium":   "#f5f9ff",
        "Low":      "#f6fbee",
        "":         "#fafafa",
    }

    # ── Build live gap dataset from uploaded questionnaire ──────────────────────
    # Gaps = items where norm is No, N/A, or Partial (and unanswered "—")
    GAP_RESPONSES = {"No", "N/A", "Partial", "—"}

    live_gaps = [i for i in p2_items if i["norm"] in GAP_RESPONSES]

    # Counts by tier × response for KPIs
    def _count(items, tier=None, resp=None):
        return sum(
            1 for i in items
            if (tier is None or i["tier"] == tier)
            and (resp is None or i["norm"] == resp)
        )

    total_gaps    = len(live_gaps)
    n_crit_total  = _count(live_gaps, tier="Critical")
    n_high_total  = _count(live_gaps, tier="High")
    n_med_total   = _count(live_gaps, tier="Medium")
    n_low_total   = _count(live_gaps, tier="Low")

    n_no_total    = _count(live_gaps, resp="No")
    n_na_total    = _count(live_gaps, resp="N/A")
    n_part_total  = _count(live_gaps, resp="Partial")
    n_unans_total = _count(live_gaps, resp="—")

    st.subheader("Gap Summary Report")
    st.caption(
        "Controls with No, N/A, or Partial responses only — sourced from the uploaded questionnaire. "
        "Filter by tier or response to prioritise remediation."
    )

    # ── KPI strip — tier breakdown ───────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total gaps",         total_gaps)
    k2.metric("🔴 Critical",        n_crit_total)
    k3.metric("🟠 High",            n_high_total)
    k4.metric("🟡 Medium",          n_med_total)
    k5.metric("🟢 Low / No tier",   n_low_total)

    # Response breakdown sub-row
    r1, r2, r3, r4, _ = st.columns([1, 1, 1, 1, 1])
    r1.markdown(
        '<div style="font-size:11px;color:#6c757d;font-weight:600;text-transform:uppercase;'
        'letter-spacing:0.05em;margin-bottom:2px">No</div>'
        f'<div style="font-size:22px;font-weight:600;color:#A32D2D">{n_no_total}</div>',
        unsafe_allow_html=True,
    )
    r2.markdown(
        '<div style="font-size:11px;color:#6c757d;font-weight:600;text-transform:uppercase;'
        'letter-spacing:0.05em;margin-bottom:2px">Partial</div>'
        f'<div style="font-size:22px;font-weight:600;color:#854F0B">{n_part_total}</div>',
        unsafe_allow_html=True,
    )
    r3.markdown(
        '<div style="font-size:11px;color:#6c757d;font-weight:600;text-transform:uppercase;'
        'letter-spacing:0.05em;margin-bottom:2px">N/A</div>'
        f'<div style="font-size:22px;font-weight:600;color:#5F5E5A">{n_na_total}</div>',
        unsafe_allow_html=True,
    )
    r4.markdown(
        '<div style="font-size:11px;color:#6c757d;font-weight:600;text-transform:uppercase;'
        'letter-spacing:0.05em;margin-bottom:2px">Unanswered</div>'
        f'<div style="font-size:22px;font-weight:600;color:#888780">{n_unans_total}</div>',
        unsafe_allow_html=True,
    )

    st.divider()

    # ── Stacked bar chart — gaps by domain × tier ────────────────────────────────
    chart_rows = []
    for letter, dom in domains.items():
        dom_gaps = [i for i in dom["items"] if i["norm"] in GAP_RESPONSES]
        for tier in ["Critical", "High", "Medium", "Low", ""]:
            cnt = sum(1 for i in dom_gaps if i["tier"] == tier)
            if cnt:
                chart_rows.append({
                    "Domain": letter,
                    "Tier": tier if tier else "No tier",
                    "Gaps": cnt,
                })

    if chart_rows:
        chart_df = pd.DataFrame(chart_rows)
        fig_gs = px.bar(
            chart_df, x="Domain", y="Gaps", color="Tier",
            color_discrete_map={
                "Critical": "#E24B4A",
                "High":     "#EF9F27",
                "Medium":   "#378ADD",
                "Low":      "#639922",
                "No tier":  "#B4B2A9",
            },
            category_orders={"Tier": ["Critical", "High", "Medium", "Low", "No tier"]},
            labels={"Gaps": "Gaps (No / N/A / Partial)", "Domain": "Domain"},
            height=260,
        )
        fig_gs.update_layout(
            margin=dict(l=0, r=0, t=10, b=0),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            legend_title_text="Risk tier",
            font=dict(size=12),
            bargap=0.3,
        )
        fig_gs.update_xaxes(showgrid=False)
        fig_gs.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.07)")
        st.plotly_chart(fig_gs, use_container_width=True)

    st.divider()

    # ── Filters ──────────────────────────────────────────────────────────────────
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        gs_tier_f = st.multiselect(
            "Filter by risk tier",
            options=["Critical", "High", "Medium", "Low", ""],
            default=["Critical", "High", "Medium", "Low", ""],
            format_func=lambda x: x if x else "No tier",
            key="gs_tier_filter",
        )
    with fc2:
        gs_resp_f = st.multiselect(
            "Filter by response",
            options=["No", "Partial", "N/A", "—"],
            default=["No", "Partial", "N/A", "—"],
            format_func=lambda x: "Unanswered" if x == "—" else x,
            key="gs_resp_filter",
        )
    with fc3:
        gs_dom_f = st.multiselect(
            "Filter by domain",
            options=sorted({i["domain_name"] for i in live_gaps}),
            default=sorted({i["domain_name"] for i in live_gaps}),
            key="gs_dom_filter",
        )

    filtered_gaps = [
        i for i in live_gaps
        if i["tier"] in gs_tier_f
        and i["norm"] in gs_resp_f
        and i["domain_name"] in gs_dom_f
    ]
    filtered_gaps.sort(key=lambda x: (TIER_ORDER.get(x["tier"], 4), x["domain"], x["key"]))

    st.caption(f"Showing {len(filtered_gaps)} of {total_gaps} gaps")
    st.divider()

    # ── Tier sections with gap cards ─────────────────────────────────────────────
    TH_GS = (
        "padding:8px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;"
        "text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef"
    )

    TIER_SECTION_LABEL = {
        "Critical": "🔴  Critical controls",
        "High":     "🟠  High controls",
        "Medium":   "🟡  Medium controls",
        "Low":      "🟢  Low controls",
        "":         "⬜  Controls without tier",
    }

    for tier_key in ["Critical", "High", "Medium", "Low", ""]:
        if tier_key not in gs_tier_f:
            continue
        tier_items = [i for i in filtered_gaps if i["tier"] == tier_key]
        if not tier_items:
            continue

        # Tier section header
        n_no   = sum(1 for i in tier_items if i["norm"] == "No")
        n_pa   = sum(1 for i in tier_items if i["norm"] == "Partial")
        n_na   = sum(1 for i in tier_items if i["norm"] == "N/A")
        n_un   = sum(1 for i in tier_items if i["norm"] == "—")
        badge  = GS_RISK_BADGE.get(tier_key, GS_RISK_BADGE.get("Low", ""))
        label  = TIER_SECTION_LABEL.get(tier_key, tier_key)

        st.markdown(
            f'<div style="display:flex;align-items:center;gap:10px;margin:8px 0 6px">'
            f'<span style="font-size:14px;font-weight:600;color:#333">{label}</span>'
            f'<span style="font-size:12px;color:#6c757d">'
            f'&nbsp;{len(tier_items)} gap{"s" if len(tier_items)!=1 else ""}'
            f'&nbsp;&nbsp;·&nbsp;&nbsp;No: <b>{n_no}</b>'
            f'&nbsp;&nbsp;Partial: <b>{n_pa}</b>'
            f'&nbsp;&nbsp;N/A: <b>{n_na}</b>'
            f'&nbsp;&nbsp;Unanswered: <b>{n_un}</b>'
            f'</span></div>',
            unsafe_allow_html=True,
        )

        # Table of gap items for this tier
        rows_html = ""
        for item in tier_items:
            resp_badge_html = f'<span style="{GS_RESP_BADGE.get(item["norm"], GS_RESP_BADGE["—"])}">{item["norm"] if item["norm"] != "—" else "Unanswered"}</span>'
            remarks_cell = (
                f'<div style="font-size:11px;color:#777;margin-top:4px;line-height:1.4">{item["other"]}</div>'
                if item["other"] else ""
            )
            rows_html += (
                "<tr>"
                f'<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:11px;'
                f'color:#aaa;width:7%;vertical-align:top;white-space:nowrap">{item["key"]}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;'
                f'color:#888;width:18%;vertical-align:top">{item["domain_name"]}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;'
                f'color:#333;line-height:1.5;width:55%;vertical-align:top">'
                f'{item["question"]}{remarks_cell}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;width:10%;'
                f'vertical-align:top;text-align:center">{resp_badge_html}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;width:10%;'
                f'vertical-align:top;text-align:center">'
                f'<span style="{GS_RISK_BADGE.get(item["tier"], GS_RISK_BADGE.get("Low",""))}">'
                f'{item["tier"] if item["tier"] else "—"}</span></td>'
                "</tr>"
            )

        tier_table_html = (
            '<div style="border:1px solid #e9ecef;border-radius:10px;overflow:hidden;margin-bottom:1.25rem">'
            '<table style="width:100%;border-collapse:collapse;table-layout:fixed">'
            '<thead><tr style="background:#f8f9fa">'
            f'<th style="{TH_GS};width:7%">Ref</th>'
            f'<th style="{TH_GS};width:18%">Domain</th>'
            f'<th style="{TH_GS};width:55%">Control / Question</th>'
            f'<th style="{TH_GS};width:10%;text-align:center">Response</th>'
            f'<th style="{TH_GS};width:10%;text-align:center">Tier</th>'
            "</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            "</table></div>"
        )
        st.markdown(tier_table_html, unsafe_allow_html=True)

    st.divider()

    # ── Domain-level recommendations (from GAP_DB) ───────────────────────────────
    active_domains = sorted({i["domain"] for i in filtered_gaps})
    if active_domains:
        st.markdown("**Recommendations by domain**")
        st.caption("Based on gaps in the filtered view above.")
        for dom_letter in active_domains:
            dom_meta = next((d for d in GAP_DB if d["id"] == dom_letter), None)
            if not dom_meta or not dom_meta.get("recs"):
                continue
            dom_gaps_count = sum(1 for i in filtered_gaps if i["domain"] == dom_letter)
            badge_html = f'<span style="{GS_RISK_BADGE.get(dom_meta["risk"], "")}">{dom_meta["risk"]}</span>'
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:8px;margin:10px 0 4px">'
                f'<span style="font-size:13px;font-weight:600;color:#333">'
                f'{dom_letter} — {dom_meta["name"]}</span>'
                f'{badge_html}'
                f'<span style="font-size:12px;color:#aaa">{dom_gaps_count} gap{"s" if dom_gaps_count!=1 else ""}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            for i, rec in enumerate(dom_meta["recs"], 1):
                st.markdown(
                    f'<div style="padding:8px 13px;border-radius:8px;background:#f8f9fa;'
                    f'border-left:3px solid #185FA5;margin-bottom:5px;font-size:12px;'
                    f'color:#333;line-height:1.55">'
                    f'<span style="font-weight:600;color:#185FA5;margin-right:6px">{i}.</span>{rec}</div>',
                    unsafe_allow_html=True,
                )

    st.divider()

    # ── Exports ──────────────────────────────────────────────────────────────────
    export_df = pd.DataFrame([{
        "Key":        i["key"],
        "Domain":     i["domain_name"],
        "Question":   i["question"],
        "Response":   i["norm"],
        "Risk Tier":  i["tier"],
        "Remarks":    i["other"],
    } for i in filtered_gaps])

    ex1, ex2, ex3 = st.columns([1, 1, 1])
    with ex1:
        st.download_button(
            "⬇ Export gaps CSV",
            data=export_df.to_csv(index=False).encode(),
            file_name="tpcra_gap_summary.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with ex2:
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Gap Summary")
        st.download_button(
            "⬇ Export gaps Excel",
            data=buf2.getvalue(),
            file_name="tpcra_gap_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with ex3:
        pdf_bytes = generate_vendor_pdf(
            GAP_DB,
            vendor_name=contact.get("vendor", ""),
            assessment_date=datetime.date.today().strftime("%d %B %Y"),
        )
        st.download_button(
            "⬇ Download vendor report PDF",
            data=pdf_bytes,
            file_name="tpcra_vendor_gap_report.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )


st.divider()
st.caption("TPCRA v3.0 — Third-Party Cyber Risk Assessment Dashboard  ·  For internal use only")
