# Tender Ops Starter

This is a small, deterministic starter for a South African tender-response workflow. It does **not** write a bid, invent bidder facts, price work, sign forms, commission affidavits, or decide legal ambiguities. It turns a source pack into four control documents:

- `source_manifest.md` — file inventory, page/sheet metadata and SHA-256 hashes.
- `requirements_register.csv` — one row per obligation, with source, owner, status and release gate.
- `red_flags.md` — contradictions, missing volumes and deadline warnings.
- `client_evidence_request.md` — evidence and approvals the bidder must provide.

The included rule set is seeded for the SC103 training pack in the parent folder. The next production step is to replace the seed with page-level extraction plus human validation for each new tender.

## Run

From the project folder:

```powershell
python .\tender_ops\tender_ops.py --source-root . --output .\tender_ops\output --as-of 2026-10-06
```

Optionally make a copy of `client_profile.example.json`, fill only verified facts, and pass it with:

```powershell
python .\tender_ops\tender_ops.py --source-root . --output .\tender_ops\output --client-profile .\tender_ops\client_profile.json
```

Run the tests with:

```powershell
python -m unittest discover -s .\tender_ops\tests -v
```

## Required human gates

1. Confirm that every referenced source volume and addendum has been received.
2. Approve the tender facts and every company fact against evidence.
3. Have an estimator/QS/engineer approve rates, outputs, risks, VAT and the final price.
4. Have authorised people sign, witness and commission the required documents.
5. Recalculate the workbook in native Excel and reconcile the offer, VAT and pricing summary.
6. Obtain a second-person compliance and packaging review before release.

Never reuse this pack's dates, forms, thresholds or submission rules for another tender. The latest issued tender documents and addenda always control.
