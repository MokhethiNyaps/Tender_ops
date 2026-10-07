#!/usr/bin/env python3
"""Generate deterministic control documents for the SC103 tender training pack."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


PACK = {
    "reference": "X.004-130-2024/1",
    "package": "SC103",
    "advertised_cidb": "3CE-4CE",
    "intention_deadline": "2026-08-07",
    "closing": "2026-08-21T12:00:00+02:00",
    "status_note": "Training pack only; confirm all facts against the latest issued documents and addenda.",
}

RETURNABLE = "Returnable Document - Subcontractor SC103.pdf"
BRIEFING = "X.004-130-2024_1 Tender Clarification Presentation 3CE-4CE SC103 Bethlehem.pdf"
WORKBOOK = "X.004-130-2024_1 SC103.xlsx"


def requirement(
    req_id: str,
    source: str,
    location: str,
    obligation: str,
    evidence: str,
    owner: str = "Bidder",
    risk: str = "HIGH",
    gate: str = "CLIENT_EVIDENCE",
    notes: str = "",
) -> dict[str, str]:
    return {
        "requirement_id": req_id,
        "source_document": source,
        "source_location": location,
        "obligation": obligation,
        "evidence_or_action": evidence,
        "owner": owner,
        "status": "OPEN",
        "risk": risk,
        "release_gate": gate,
        "notes": notes,
    }


REQUIREMENTS = [
    requirement("SOURCE-01", BRIEFING, "p.8", "Obtain every referenced tender volume and addendum before drafting.", "Tender Data/T1, C1.2.1, C3, C4, C5, drawings, C1.3.1 and all addenda.", "Tender lead", "FATAL", "SOURCE_COMPLETE"),
    requirement("A1", RETURNABLE, "p.8", "Provide the compulsory briefing attendance certificate.", "Form signed by bidder and employer representative.", gate="SIGNATURE"),
    requirement("A1.1", RETURNABLE, "p.9", "Submit the intention-to-tender certificate by the stated deadline.", "Signed certificate plus email transmission proof.", gate="DEADLINE"),
    requirement("A2.1", RETURNABLE, "p.10", "Prove authority of the tender signatory.", "Actual signed resolution on company letterhead; do not rely on the sample page alone.", gate="SIGNATURE"),
    requirement("A2.2", RETURNABLE, "p.11", "Declare that the entity participates in only one tender.", "Signed declaration for bidder and each applicable JV member/targeted enterprise.", gate="SIGNATURE"),
    requirement("A2.3", RETURNABLE, "pp.12-13", "Complete the fronting-practices declaration truthfully.", "Client-approved and signed disclosure.", gate="SIGNATURE"),
    requirement("A2.4", RETURNABLE, "pp.14-16", "Complete DPIP/FPPO/FIN and ownership declarations.", "Verified ownership and identity details; separate declarations where required.", gate="CLIENT_FACT_APPROVAL"),
    requirement("A2.5", RETURNABLE, "p.17", "Provide consent to due-diligence investigation.", "Signed consent.", gate="SIGNATURE"),
    requirement("A3.1", RETURNABLE, "pp.18-21", "Complete bidder disclosure, conflicts and related-interest declarations.", "Client-approved SBD4-style disclosure and signatures.", gate="SIGNATURE"),
    requirement("A3.2", RETURNABLE, "pp.22-23", "Resolve and, unless removed by addendum, complete the independent-tender certificate.", "Written clarification because the form is marked REPEALED but remains in the mandatory list.", "Tender lead", "HIGH", "CLARIFICATION"),
    requirement("A3.3", RETURNABLE, "pp.24-25", "Declare past SCM practices and restriction history.", "Verified client disclosure and signature.", gate="SIGNATURE"),
    requirement("A3.4", RETURNABLE, "p.26", "Provide CSD registration evidence current at closing.", "CSD report and supplier number for bidder and each JV member.", gate="ELIGIBILITY"),
    requirement("A3.5", RETURNABLE, "pp.27-31", "Complete the local-production/content declaration where applicable.", "Approved SBD6.2 response plus supporting calculations and responsible-person signature.", gate="SPECIALIST_REVIEW"),
    requirement("A3.6", RETURNABLE, "p.32", "Submit the local-content Annex C summary.", "Completed Annex C reconciled to retained supporting schedules.", gate="SPECIALIST_REVIEW"),
    requirement("A4", RETURNABLE, "p.36", "List every deviation or state none.", "Client/legal approval and signed schedule.", gate="CLIENT_FACT_APPROVAL"),
    requirement("A5", RETURNABLE, "p.37", "Record every addendum received.", "Addenda register reconciled to source manifest.", "Tender lead", "FATAL", "SOURCE_COMPLETE"),
    requirement("A6", RETURNABLE, "p.38", "Provide tax-compliance consent and PIN details.", "SARS tax PIN/reference authorised by bidder.", gate="ELIGIBILITY"),
    requirement("A7", RETURNABLE, "pp.39-41", "Provide a recent bank/account-conduct confirmation.", "Bank letter less than three months old; resolve compiler notes/options.", gate="CLIENT_EVIDENCE"),
    requirement("A8", RETURNABLE, "p.42", "Disclose current contracts and unawarded tenders.", "Verified commitments schedule.", gate="CLIENT_FACT_APPROVAL"),
    requirement("A9", RETURNABLE, "p.43", "Provide COID good standing and labour-inspection disclosures.", "Valid COID evidence and verified 36-month disclosure.", gate="ELIGIBILITY"),
    requirement("A10", RETURNABLE, "p.44", "Prove the required CIDB status.", "Current CIDB listing and renewal proof if required.", gate="ELIGIBILITY"),
    requirement("A11", RETURNABLE, "pp.45-46", "Provide the correct B-BBEE/EME/QSE evidence for the bidder's verified facts.", "Valid certificate or legally suitable sworn evidence; never select an affidavit from grade alone.", gate="ELIGIBILITY"),
    requirement("B1", RETURNABLE, "p.47", "Provide relevant work experience.", "Evidence-backed schedule with employer/engineer contacts, value and completion year.", gate="CLIENT_EVIDENCE"),
    requirement("B2", RETURNABLE, "p.48", "Provide the equipment schedule.", "Verified owned/on-order/hired equipment and supporting proof.", gate="CLIENT_EVIDENCE"),
    requirement("C1", RETURNABLE, "p.49", "Complete establishment/general-obligations percentages within stated caps.", "Estimator-approved calculation; explanatory letter if the combined total exceeds 15%.", "Estimator", "FATAL", "PRICE_APPROVAL"),
    requirement("C1.1.1", RETURNABLE, "pp.50-51", "Complete and sign the Form of Offer with the VAT-inclusive total.", "Native-workbook total reconciled to the offer; authorised signature and witnesses.", "Director/signatory", "FATAL", "SIGNATURE"),
    requirement("C1.1.2", RETURNABLE, "pp.52-53", "Do not pre-complete the employer's Form of Acceptance as a bidder response unless instructed.", "Post-award/employer control.", "Employer", "MEDIUM", "POST_AWARD"),
    requirement("C1.2.3", RETURNABLE, "p.54", "Provide tenderer communications, representative and director details.", "Verified company facts approved by client.", gate="CLIENT_FACT_APPROVAL"),
    requirement("C2.2", WORKBOOK, "P's& G's E10:F58; C1.5 (2) G9:I23", "Complete the pricing schedule without altering protected quantities or formulas.", "Estimator-approved rates; native Excel recalculation; scanned and editable copies reconciled.", "Estimator", "FATAL", "PRICE_APPROVAL"),
    requirement("D1", RETURNABLE, "p.56", "Complete and sign the final compliance checklist.", "Independent readback, ticks, signature and package reconciliation.", "Independent reviewer", "FATAL", "FINAL_QA"),
    requirement("SUBMISSION-01", BRIEFING, "pp.23,26-27", "Confirm electronic media, delivery address, envelope marking and copy requirements in writing.", "Formal clarification/addendum because the pack conflicts on CD/flash drive and gives two addresses.", "Tender lead", "FATAL", "CLARIFICATION"),
]


RED_FLAGS = [
    ("FATAL", "The pack closed on 21 August 2026 at 12:00 and is expired for live submission.", f"{BRIEFING} p.27", "Use only as a training case; obtain the client's current live tender."),
    ("FATAL", "Referenced source volumes are missing: T1/Tender Data, C1.2.1, C3, C4, C5, drawings and C1.3.1.", f"{BRIEFING} p.8; {RETURNABLE} pp.52-53", "Request the complete issued pack and all addenda."),
    ("HIGH", "The main-project title names R26/R34/R74 while the package scope identifies R57.", f"{RETURNABLE} p.2; {BRIEFING} pp.2,5", "Submit a written clarification; do not silently choose."),
    ("HIGH", "The functionality threshold is shown only as an example, and the applicable 3CE/4CE scoring column is unclear.", f"{BRIEFING} pp.9-11", "Obtain the confirmed threshold, scoring column and symbol legend."),
    ("HIGH", "The listed functionality components can appear to total 105 before any normalization.", f"{BRIEFING} p.9", "Obtain the employer's confirmed scoring method."),
    ("HIGH", "The returnable says flash drive; the presentation says compact disc and also requires scanned plus Excel pricing.", f"{RETURNABLE} pp.6,56; {BRIEFING} pp.23,26", "Confirm required media and file formats in writing."),
    ("HIGH", "Two hand-delivery addresses are joined by 'or' without a package allocation.", f"{BRIEFING} p.27", "Confirm the correct tender box and address."),
    ("HIGH", "A3.2 is marked REPEALED but remains on the mandatory lists.", f"{RETURNABLE} pp.6,22-23,56", "Complete unless a written addendum removes it; request clarification."),
    ("HIGH", "The CSD-at-closing rule conflicts with a presentation statement about a five-day cure.", f"{RETURNABLE} p.26; {BRIEFING} pp.12-13", "Treat at-closing evidence as mandatory unless formally clarified."),
    ("HIGH", "The pack appears to use old SBD6.1/SBD6.2 preference/local-content material.", f"{RETURNABLE} pp.27-32", "Do not delete issued forms; request a formal clarification/addendum on the applicable 2022 rules."),
    ("MEDIUM", "A7 includes unresolved compiler notes/options and the returnable has broken bookmarks/placeholders.", f"{RETURNABLE} pp.2,7,39-41", "Preserve the source and obtain instructions before removing or choosing text."),
    ("HIGH", "The supplied QSE affidavit is for a specialised entity and is not a generic construction-contractor QSE affidavit.", "Affidavit-QSE-Spec.pdf pp.1-2", "Choose B-BBEE evidence only from verified entity type, turnover and ownership facts."),
    ("HIGH", "Pricing input cells are not visibly identified; two unlocked workbook cells have no unit/quantity.", f"{WORKBOOK}: C1.5 (2)!H16:H17", "Have the issuer/estimator confirm applicability before entering rates."),
]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def pdf_page_count(path: Path) -> int | None:
    data = path.read_bytes()
    matches = re.findall(rb"/Type\s*/Page\b", data)
    return len(matches) or None


def xlsx_sheet_names(path: Path) -> list[str]:
    try:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError):
        return []
    namespace = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return [node.attrib.get("name", "") for node in root.findall(".//m:sheet", namespace)]


def load_profile(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def profile_value(profile: dict[str, Any], dotted_key: str) -> Any:
    value: Any = profile
    for part in dotted_key.split("."):
        if not isinstance(value, dict):
            return None
        value = value.get(part)
    return value


def inventory(source_root: Path, output_root: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    supported = {".pdf", ".xlsx", ".xls", ".docx", ".doc", ".csv"}
    for path in sorted(source_root.iterdir(), key=lambda p: p.name.casefold()):
        if not path.is_file() or path.suffix.lower() not in supported:
            continue
        if output_root in path.parents:
            continue
        detail = ""
        if path.suffix.lower() == ".pdf":
            count = pdf_page_count(path)
            detail = f"{count} pages" if count is not None else "page count unavailable"
        elif path.suffix.lower() == ".xlsx":
            sheets = xlsx_sheet_names(path)
            detail = "sheets: " + ", ".join(sheets) if sheets else "sheet metadata unavailable"
        rows.append({
            "filename": path.name,
            "type": path.suffix.lower().lstrip("."),
            "bytes": path.stat().st_size,
            "detail": detail,
            "sha256": sha256_file(path),
        })
    return rows


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def manifest_markdown(rows: Iterable[dict[str, Any]], as_of: date) -> str:
    closing = datetime.fromisoformat(PACK["closing"])
    status = "EXPIRED" if as_of > closing.date() else "OPEN OR UNCONFIRMED"
    lines = [
        "# SC103 source manifest",
        "",
        f"Generated: {as_of.isoformat()}",
        f"Tender reference: {PACK['reference']} / {PACK['package']}",
        f"Pack status: **{status}**",
        "",
        "Hashes identify the exact files reviewed. Any changed or added file requires a fresh run and requirement review.",
        "",
        "| File | Type | Size (bytes) | Metadata | SHA-256 |",
        "|---|---:|---:|---|---|",
    ]
    for row in rows:
        lines.append(f"| {row['filename']} | {row['type']} | {row['bytes']} | {row['detail']} | `{row['sha256']}` |")
    lines.extend(["", PACK["status_note"], ""])
    return "\n".join(lines)


def client_request_markdown(profile: dict[str, Any]) -> str:
    checks = [
        ("Complete issued tender pack, drawings and every addendum", None),
        ("CIPC/company registration and verified legal name", "entity.legal_name"),
        ("Signed signatory-authority resolution on letterhead", "entity.signatory_authority_evidence"),
        ("Current CSD report and supplier number", "registrations.csd_supplier_number"),
        ("SARS tax-compliance PIN", "registrations.sars_tax_pin"),
        ("CIDB CRS number and current grade/class", "registrations.cidb_grade_class"),
        ("COID letter of good standing", "registrations.coid_good_standing_expiry"),
        ("Correct B-BBEE certificate or affidavit supported by verified entity facts", "bbbee.evidence_type"),
        ("Bank/account-conduct letter dated within the tender window", "commercial.bank_letter_date"),
        ("Locality proof", "capability.locality_evidence"),
        ("Project references and completion evidence", "capability.completed_project_references"),
        ("Personnel CVs, qualifications and availability", "capability.key_personnel"),
        ("Equipment ownership/hire evidence", "capability.equipment"),
        ("Current commitments and unawarded tenders", "capability.current_commitments"),
        ("Supplier quotations, productivity assumptions, risks, margin and VAT decisions", None),
        ("Written authority for any portal/account access or submission action", None),
    ]
    lines = [
        "# Client evidence and approval request",
        "",
        "Provide only current, authentic records. `OPEN` means the profile does not contain a verified value; it is not permission to infer one.",
        "",
        "| Status | Required item |",
        "|---|---|",
    ]
    for label, key in checks:
        value = profile_value(profile, key) if key else None
        supplied = bool(value)
        lines.append(f"| {'SUPPLIED — VERIFY' if supplied else 'OPEN'} | {label} |")
    lines.extend([
        "",
        "## Approvals that cannot be automated",
        "",
        "- Client approval of every company fact, disclosure, deviation and claim.",
        "- Estimator/QS/engineer approval of rates, outputs, caps, VAT and final price.",
        "- Authorised signatures, witnesses and commissioning of sworn documents.",
        "- Independent final compliance, scan, workbook and packaging readback.",
        "",
    ])
    return "\n".join(lines)


def red_flags_markdown(as_of: date) -> str:
    lines = [
        "# SC103 red flags and clarification log",
        "",
        f"Review date: {as_of.isoformat()}",
        "",
        "Do not resolve conflicts by assumption. Record an issuer response or addendum and link it to every affected requirement.",
        "",
        "| Severity | Issue | Source | Required action | Status |",
        "|---|---|---|---|---|",
    ]
    for severity, issue, source, action in RED_FLAGS:
        lines.append(f"| {severity} | {issue} | {source} | {action} | OPEN |")
    lines.append("")
    return "\n".join(lines)


def write_requirements_csv(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = list(REQUIREMENTS[0].keys())
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(REQUIREMENTS)


def generate(source_root: Path, output_root: Path, as_of: date, profile_path: Path | None = None) -> dict[str, Any]:
    source_root = source_root.resolve()
    output_root = output_root.resolve()
    profile = load_profile(profile_path)
    files = inventory(source_root, output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    write_text(output_root / "source_manifest.md", manifest_markdown(files, as_of))
    write_requirements_csv(output_root / "requirements_register.csv")
    write_text(output_root / "red_flags.md", red_flags_markdown(as_of))
    write_text(output_root / "client_evidence_request.md", client_request_markdown(profile))
    summary = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "as_of": as_of.isoformat(),
        "tender_reference": PACK["reference"],
        "package": PACK["package"],
        "source_files": len(files),
        "requirements": len(REQUIREMENTS),
        "red_flags": len(RED_FLAGS),
        "submission_ready": False,
        "reason": "Human evidence, clarifications, pricing approvals, signatures and final QA remain open.",
    }
    write_text(output_root / "run_summary.json", json.dumps(summary, indent=2) + "\n")
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "output")
    parser.add_argument("--as-of", type=date.fromisoformat, default=date.today())
    parser.add_argument("--client-profile", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = generate(args.source_root, args.output, args.as_of, args.client_profile)
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
