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
│   └── report_viewer.py     # HTML/XLSX report generator
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

## Deployment (Azure App Service)

### Deploy — Kudu File Manager (single file update)

1. `portal.azure.com → <app> → Development Tools → Advanced Tools → Go`
2. **File Manager** → `home/site/wwwroot/`
3. Drag and drop the modified file
4. Restart the app from the portal or via CLI:

```bash
az webapp restart --name <app-name> --resource-group <resource-group>
```

### Deploy — Full ZIP (Azure CLI)

```powershell
$APP = "<app-name>"
$RG  = "<resource-group>"

az webapp stop   --name $APP --resource-group $RG
az webapp deploy --name $APP --resource-group $RG `
    --src-path "$env:USERPROFILE\Downloads\pillar_deploy.zip" --type zip
az webapp start  --name $APP --resource-group $RG
```

### ZIP contents

```
app.py
requirements.txt
globe_codes.json
services/__init__.py
services/xml_validator.py
services/shell_telematico.py
services/report_viewer.py
templates/index.html
```

### Useful CLI commands

```bash
# Live log streaming
az webapp log tail --name <app-name> --resource-group <resource-group>

# Check app status
az webapp show --name <app-name> --resource-group <resource-group> \
    --query "{stato:state}" --output table

# Stop the app
az webapp stop --name <app-name> --resource-group <resource-group>
```

---

## Authentication

Authentication is implemented inline in `app.py` using **MSAL** (Microsoft Authentication Library).

**Flow:** open site → automatic Microsoft redirect → login → `/auth/callback` → home

- Accepts any **organizational Microsoft account**
- Blocks personal Microsoft accounts (`tid = 9188040d-...`)
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

---

## Validator Changelog

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
- The job store (`JOBS` dict in `app.py`) is in-memory and does not survive app restarts. For persistent job history, wire up Azure Table Storage or a database.
- Multiple XML files can be processed in a single upload, each with its own *Codice Fiscale* fornitore (`cf_map[]` parallel array in the form POST).