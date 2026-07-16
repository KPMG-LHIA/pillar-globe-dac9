# PILLAR — GloBE/DAC9 Transmission Pipeline

**PILLAR** is a web application that automates the **GIR26** (Pillar Two / DAC9 / GloBE Information Return) transmission pipeline to the *Agenzia delle Entrate* (AdE) via Entratel.

---

## Overview

PILLAR accepts a GloBE XML file in the standard OECD format (`globe:GLOBE_OECD v2`) and runs it through a three-step pipeline:

```
Input: XML GloBE (OECD GLOBE_OECD v2 format)
  │
  ├── Step 1 — xml_validator.py      →  <stem>_validation_report.xlsx
  ├── Step 2 — shell_telematico.py   →  <stem>_MSG.xml
  └── Step 3 — report_viewer.py      →  <stem>_report.html
```

| Step | Module | Output | Description |
|------|--------|--------|-------------|
| 1 | `xml_validator.py` | `.xlsx` | Validates the XML against 155+ AdE Allegato 2 checks |
| 2 | `shell_telematico.py` | `_MSG.xml` | Encapsulates the XML in the AdE *shell telematica* (Entratel format) |
| 3 | `report_viewer.py` | `.html` | Generates a navigable HTML report with jurisdiction-level detail |

---

## Regulatory References

| Document | Description |
|----------|-------------|
| Provvedimento AdE Prot. **112451/2026** | Italian implementing provision for DAC9/GIR26 |
| **Allegato 2** AdE — 13/03/2026 | Technical control rules (§8.1, §8.2, §8.3) |
| Interpretive rules **v1.1** — 28/05/2026 | AdE clarifications on control checks |
| `GLOBEXML_v1.0.xsd` | OECD GloBE Information Return XML schema |
| `fornituraGIRT_v1.0.0.xsd` | AdE GIR transmission schema |
| `telematico_v1.xsd` | AdE Entratel shell schema |
| OECD GloBE STF `globestf_v5` | Standard Transmission Format XSD |

---

## Security & Reliability Hardening (2026-05/07)

| Area | Before | After |
|------|--------|-------|
| Job/file persistence | In-memory (`JOBS = {}`), lost on every App Service restart | Azure Blob Storage (`services/storage.py`, containers `pillar-files`/`pillar-jobs`), local filesystem fallback for dev |
| Secrets | Env vars only | Azure Key Vault + System-assigned Managed Identity (RBAC) for `AzureAdClientSecret`, `FlaskSecretKey`, `StorageConnectionString` |
| `FLASK_SECRET_KEY` | Regenerated at every restart → sessions invalidated | Persisted as an App Setting (from Key Vault) |
| Logging | Ad-hoc lists in RAM | Structured `logging` module output, collected by App Service log / App Insights |
| Storage encryption/DR | N/A | Blob Storage SSE at rest; RA-GRS (North Europe → West Europe) |
| Background job resilience | Import failures inside `run_pipeline` could leave a job stuck at `PENDING`/`RUNNING` forever | Imports moved inside `try/except`; any failure now marks the job `FAILED` with a visible error message |

---

## Project Structure

```
PILLAR/
├── app.py                   # Flask app — MSAL OAuth2 auth, async pipeline worker
├── requirements.txt         # Python dependencies
├── globe_codes.json         # GIR code dictionary (ETRRange, SafeHarbour, DocType, …)
├── services/
│   ├── __init__.py
│   ├── xml_validator.py     # 155+ AdE checks (Allegato 2, §8.1–8.3)
│   ├── shell_telematico.py  # Shell MSG builder + namespace normalisation
│   ├── report_viewer.py     # HTML/XLSX report generator
│   └── storage.py           # Persistent storage abstraction (Azure Blob Storage,
│                             #   containers `pillar-files` / `pillar-jobs`, with
│                             #   local filesystem fallback for dev) — used for
│                             #   uploaded files, pipeline outputs, and job state
└── templates/
    └── index.html           # UI — multi-TIN upload, live log, grouped downloads
```

---

## Requirements

- Python 3.11
- Dependencies (see `requirements.txt`):

```
flask>=3.0.0
lxml>=5.0.0
openpyxl>=3.1.0
msal>=1.28.0
gunicorn>=21.0.0
```

> **Nota:** `services/storage.py` usa Azure Blob Storage (System-assigned Managed
> Identity + Key Vault) in produzione, con fallback su filesystem locale in
> assenza di credenziali Azure. Se non già presenti, verificare che
> `requirements.txt` includa `azure-storage-blob` e `azure-identity`.

---

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run locally — leave AZURE_AD_CLIENT_ID unset to bypass authentication
python app.py
# → http://127.0.0.1:5000
```

When `AZURE_AD_CLIENT_ID` is not set, the app auto-creates a local dev session (`Dev Locale / dev@localhost`) and skips the Microsoft login flow entirely.

### Running the validator standalone

```bash
python services/xml_validator.py path/to/file.xml [output_dir]
# → produces path/to/file_validation_report.xlsx
```

---

## Authentication

Authentication is implemented inline in `app.py` using **MSAL** (Microsoft Authentication Library).

**Flow:** open site → automatic Microsoft redirect → login → `/auth/callback` → home

- Accepts any **organizational Microsoft account**
- Blocks personal Microsoft accounts
- Authority: `https://login.microsoftonline.com/<tenant-id>` (single-tenant)
- Access control is delegated to Azure AD (Enterprise Applications → App Registration)

The authority is tenant-specific, derived from the `AZURE_TENANT_ID` environment variable set on the App Service. To allow users from external tenants, admin consent must be granted on the target tenant.

---

## Validator — Check Coverage

`xml_validator.py` implements all 155+ control rules verifiable offline from AdE Allegato 2 (13/03/2026).

| Category | Range | Total | Implemented | Skipped (offline) |
|----------|-------|-------|-------------|-------------------|
| File integrity | C0000–C0004 | 3 | 3 | 0 |
| File errors | 50001–50004 | 4 | 4 | 0 |
| Record errors — Severe | 60001–60028 | 28 | 24 | 4 |
| Record errors — Other | 70001–70124 | 124 | 124 | 0 |
| **Total** | | **159** | **155** | **4** |

### Checks skipped (require AdE transmission history)

| Code | Reason |
|------|--------|
| 60002 | `MessageRefId` duplicate — requires AdE transmission history |
| 60008 | `CorrDocRefId` references unknown record — requires AdE history |
| 60009 | `CorrDocRefId` must point to last valid version — requires AdE history |
| 60014 | `DocRefId` OECD0 matches last version — requires AdE history |

### Validation report (XLSX output)

| Sheet | Contents |
|-------|----------|
| **Sommario** | Overall status, counters, transmission verdict |
| **Dettaglio** | All 155+ checks with OK / KO / SKIP / WARN status and XML line numbers |
| **Errori e Warning** | KO and WARN rows only, with full detail |

---

## Known Issues & Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| `S0001` from AdE | Source XML has `p3:version` instead of plain `version="1.0"` on `GLOBE_OECD` | Fixed automatically by `shell_telematico.py` |
| `S0001` from AdE | Inline namespace redeclarations on `DocTypeIndic` / `DocRefId` | Fixed automatically by `shell_telematico.py` via `etree.cleanup_namespaces` |
| `S0001` from AdE | Entirely wrong `JurisdictionSection` structure (invented elements, missing required GloBE data) | Source XML error — must be fixed by the client's software vendor |
| `{Guid:D}` placeholders | Unresolved GUIDs in source XML | Resolved in-memory by both validator and shell; source file is never modified |
| UI stuck on "loading" indefinitely after a deploy | `run_pipeline`'s module imports (`xml_validator`, `shell_telematico`, `report_viewer`) sat outside the `try/except` — an import failure (e.g. from a non-atomic Kudu upload) killed the background thread silently, leaving the job at `PENDING`/`RUNNING` forever | Fixed (2026-07): imports moved inside `try`; any failure, including at import time, now marks the job `FAILED` with a visible error |
| Report showed `DocTypeIndic=OECD1` as "Resent data" | `globe_codes.json`'s `TYPE_INDIC` map had OECD0/OECD1 (and OECD10-13) inverted relative to the official XSD enum `6.4.33 STF:OECDDocTypeIndic_EnumType` | Fixed (2026-07): `TYPE_INDIC` corrected — `OECD0`=Resend Data, `OECD1`=New Data, `OECD2`=Corrected Data, `OECD3`=Deletion of Data (+ `OECD10-13` Test variants) |

---

## Validator Changelog

### v1.3.4 (2026-07)
Fixes cross-checked against an official AdE validation report for a real transmission file (`IT_07224860960.xml`):
- **70033** — severity fix: `Exception/TIN` pointing to the UPE (OtherUPE/ExcludedUPE) is `W7033` (Warning) per AdE, not an error. Confirmed on 18/18 occurrences in the AdE report; only "TIN matches neither a CE nor the UPE" remains KO
- **70079** — tolerance fix (was `>1`, silently hiding a 1-unit rounding mismatch that AdE does flag) + severity fix (KO → WARN, matching `W7079`)
- **70086** — same tolerance + severity fix as 70079, matching `W7086`
- **70120** — structural bug fix: `Recast/Higher` and `Recast/Lower` are nested inside each `Adjustment` element, not a direct child of `DeferTaxAdjustAmt` as the old XPath assumed; the old code always read `Recast` as `0`, producing false-positive KOs on every file with a non-zero Recast
- **70121** — severity fix: KO → WARN, matching `W7121` (the `bad_warn` list existed in the code since v1.2 but was never actually populated/used — dead code)
- **70122** — logic fix: opposite-sign check was comparing *any* two `Adjustment/Amount` values regardless of `AdjustmentItem` code (and regardless of whether one of them was legitimately `0`, e.g. an `Adjustment` that only carries a `Recast`); now scoped to pairs sharing the same `AdjustmentItem` code, mirroring the 70114/70118 fix from v1.2
- **`_warn_all()`** — new helper (mirrors `_ko_all()` but sets `STATUS_WARN`), used by the checks above

### v1.3.1 (2026-06)
- **70026** — fix: `OwnershipPercentage` accepts both fraction (`1`) and percentage (`100`) notation; both are valid per AdE spec
- **lxml** — fix: replaced all element truth-tests (`if el:`) with `if el is not None:`; an lxml element with no children evaluates to `False`, silently skipping checks

### v1.3 (2026-06)
- **70004** — fix: excludes `TypeOfTIN=GIR3002` from the `issuedBy == ResCountryCode` check (GIR3002 = TIN issued by a different jurisdiction for internal registration; `issuedBy` may legitimately differ from `ResCountryCode`)
- **70016** — fix: XML element name is `GlobeStatus` (not `GloBEStatus`); use `_findall` to collect multiple values per ID node
- **`_gs()`** — new helper replacing all occurrences of `set((_t(_find(...GloBEStatus...)) or "").split())`

### v1.2 (2026-06)
- **`_src()`** — helper adding `(riga XML: N)` to KO detail messages
- **70033** — fix: TIN lookup includes OtherUPE/ExcludedUPE
- **70034/70035** — fix: check Art2.1.3/Art2.1.5 Status only when element is present
- **70047/70048** — fix: SafeHarbour→ETR mapping by SubGroup TIN, not by jurisdiction
- **70114/70118** — fix: opposite-sign check scoped to same AdjustmentItem code
- **70121** — fix: correct scope = per single CEComputation, not aggregated ETR

---

## Architecture Notes

- The pipeline runs all three steps sequentially regardless of validation outcome — errors in Step 1 do not block Steps 2 and 3.
- `{Guid:D}` placeholders are resolved in-memory at runtime; the source file on disk is never modified.
- Checks 60002/60008/60009/60014 require AdE transmission history and cannot be verified offline — they are permanently marked SKIP.
- The job store is persisted via `services/storage.py` (Azure Blob Storage, container `pillar-jobs`, with local filesystem fallback for dev) and survives App Service restarts — this superseded the earlier in-memory `JOBS` dict, which did not survive restarts (frequent on the F1 plan).
- Uploaded files and pipeline outputs are likewise persisted via `services/storage.py` (container `pillar-files`), downloaded to a temporary local directory for processing (lxml/openpyxl require filesystem access) and cleaned up at the end of each job.
- Multiple XML files can be processed in a single upload, each with its own *Codice Fiscale* fornitore (`cf_map[]` parallel array in the form POST).
- `report_viewer.py` surfaces `RecJurCode` (both at `FilingInfo`/`GeneralSection` level and per-`Summary`) and, in the Corporate Structure table, the `QIIR` block (`POPE-IPE` code, which `ExceptionRule` article is `true`, and the referenced `TIN`) — previously present in the source XML but not shown in the HTML report.