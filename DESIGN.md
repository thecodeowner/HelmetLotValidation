# Helmet Lot Validation — Design Document

> Status: **Planning / draft.** Built from the requirements discussion and the sample
> file `Lot_LG01_SN_Galvion_Selection_Clean.xlsx`. Items marked **[TO INVESTIGATE]**
> still need confirmation before/while building.

## 1. Purpose

A standalone Windows desktop application that scans a **production batch (lot)
configuration file** for a product and performs three jobs:

1. **Sanity check** the file against 3–4 Microsoft Access databases (Molding,
   Assembly, Quality, Bonding — which ones apply depends on the product).
2. **Gather the physical documentation** for every sub-component lot into a
   structured **"hand-over package"**, pull the product's full Bill of Material
   from the Arena PLM web API, and validate each document (PO + Certificate of
   Conformity) against the lot file and the BOM.
3. **Generate a comprehensive, colour-coded report** per serial number stating
   what is confirmed good, confirmed wrong, or unknown/needs manual check.

The app runs **fully on a Windows PC**, reads from **network shares** and Access
DBs (read-only) and the **Arena REST API**, and writes the report + package to disk.

## 2. Recommended technology stack

**Recommendation: Python, packaged into a single Windows `.exe`.**

| Concern | Choice | Why |
|---|---|---|
| Language/runtime | Python 3.12 | Fast to build/iterate; strong libraries for every piece below |
| GUI | PySide6 (Qt) | Native-looking desktop window, real progress/status bar, background worker threads so the UI never freezes during long runs |
| Access DB access | `pyodbc` via the **Microsoft Access Database Engine** ODBC driver | Reliable read of `.mdb`/`.accdb` on the target PC |
| Excel read | `openpyxl` | Reads the lot file; also writes the coloured report |
| Report write | `openpyxl` (xlsx) + optional PDF snapshot | Colour cell fills mirror the lot layout |
| PDF / OCR | `pypdf` (native text) + `pdf2image` + **Tesseract** OCR (scanned pages) | Certs/POs are scanned images → OCR required |
| Arena API | `requests` | Simple REST calls + session token header |
| Packaging | **PyInstaller** one-file `.exe` | Standalone, no Python install needed on the PC |

**IT prerequisites on the target PC** (to be documented in a README):
- Microsoft Access Database Engine (ODBC) — matching bitness (build 64-bit).
- Tesseract OCR + Poppler (for `pdf2image`), bundled or installed alongside.
- Read access to the `M:` share and the Access DBs; outbound HTTPS to Arena.

> If a polished MSI installer and zero external runtime are hard requirements for
> IT sign-off, a C#/.NET WPF port is the alternative; noted but not recommended
> for speed of first delivery.

## 3. High-level architecture

```
+-------------------- GUI (PySide6) --------------------+
|  Pick lot file | Choose DBs/product | Run | Progress  |
+----------------------------+--------------------------+
                             | runs on a worker thread
        +--------------------v---------------------+
        |            Orchestrator                  |
        |  loads file -> runs phases -> report     |
        +--+-------------+-------------+-----------+
           |             |             |
   +-------v----+  +-----v------+ +----v---------+
   | LotFile    |  | DB Sanity  | | Document /   |
   | Parser     |  | Checker    | | Package      |
   | (xlsx)     |  | (pyodbc)   | | Builder      |
   +------------+  +------------+ +----+---------+
                                      |
                        +-------------+-------------+
                        | Arena API | Tree Search | OCR |
                        +---------------------------+
                             |
                     +-------v--------+
                     | Report Writer  |  -> hand-over package root
                     +----------------+
```

Modules (Python packages):
- `lotfile/` — parse the workbook into typed rows (see §4).
- `lots/` — the messy-lot-string parser (see §7). Shared by DB + document phases.
- `db/` — one adapter per Access DB; a product→DB mapping resolver.
- `arena/` — login, item lookup, recursive BOM fetch, in-memory cache.
- `docs/` — supplier-tree search, PDF/OCR text extraction, PO & part-number matching.
- `report/` — summary sheet + coloured detail sheet writer.
- `config/` — load/validate the external config file; expose typed settings to all modules.
- `app/` — GUI, orchestrator, status/progress plumbing.

## 3.1 Configuration file

Nothing environment-specific is hard-coded. All credentials and anything that
can change between PCs, sites, or over time live in an **external editable config
file** next to the `.exe`, loaded at startup and validated (with clear errors if
something is missing or malformed).

**Format:** a human-readable `config.yaml` (or `.ini`), plus a first-run
GUI **Settings** screen so a non-technical operator can edit it without a text
editor.

**Location & precedence:** the app looks for the config in order: (1) a path given
on the command line / an env var, (2) `%PROGRAMDATA%\HelmetLotValidation\config.yaml`
(shared, machine-wide), (3) `config.yaml` beside the `.exe`. A bundled
`config.example.yaml` documents every key.

**Contents (illustrative):**
```yaml
databases:
  molding:   { path: "\\\\server\\share\\Molding.accdb",  password: "" }
  assembly:  { path: "\\\\server\\share\\Assembly.accdb", password: "" }
  quality:   { path: "\\\\server\\share\\Quality.accdb",  password: "" }
  bonding:   { path: "\\\\server\\share\\Bonding.accdb",  password: "" }

paths:
  supplier_root: "M:\\Armor\\Newport\\QUALITY\\Incoming_Inspection\\Supplier"
  output_root:   "M:\\Armor\\Newport\\QUALITY\\HandOver"   # where packages/reports go

arena:
  base_url:     "https://api.arenasolutions.com/v1"
  email:        "integration.user@company.com"
  password:     ""              # see credential handling below
  workspace_id: 123456789
  bom_max_depth: 0              # 0 = full explosion

matching:
  po_prefixes: ["PR", "PE", "PC"]
  weight_tolerance_source: "base_information"   # tolerances read from the file
  text_compare: { trim: true, case_insensitive: true }

product_db_map:                 # which DBs apply per product/part number (see §5, A4)
  "4-8762-0039": [molding, assembly, quality, bonding]

ocr:
  tesseract_path: "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
  poppler_path:   "C:\\Tools\\poppler\\bin"
```

**Credential handling.** Storing passwords in plain text on a shared PC is a risk.
Options, in order of preference:
1. **Windows DPAPI-encrypted** secrets (via the Settings screen), decryptable only
   by the same Windows user/machine — passwords never sit in plain text on disk.
2. **Windows Credential Manager** entries referenced by name from the config.
3. Plain text in the config as a last resort (documented as not recommended).
Arena SSO note: a **dedicated integration user** is recommended over a personal login.

> **[TO INVESTIGATE]** Whether the DBs/Arena use shared service credentials or the
> logged-in user's Windows auth, and which secret-storage option IT will accept.

## 4. Input file specification (from the sample)

Three sheets:

- **Base information** (key/value): `Product Part number`, `Min molded Weight`,
  `Max molded Weight`, `Min assembled Weight`, `Max assembled weight`.
- **Serial Number List**: a size header row (e.g. "Large") then one serial per row.
- **`LOT <id> Traceability`**: one row per serial, 42 columns. Header names come
  from row 1. Notable columns:
  - Identity: `SerialNumber`, `Mold-ID`, `Item` (part number), `Size`, `Product`,
    `sItemDescription`, `sColour`, `Shipment`, `ShippingPallet`.
  - Weights: `Weight` (molded), `iCombatWeight` (assembled).
  - Dates: `dtInspected`, `dtAssembled` (some values originate from the DBs).
  - Sub-component lots (gather docs for these): `Velcro Lot - Ext.`,
    `Velcro Lot - Int.`, `sHarnessLot`, `sBoltLot`, `sRearImpactPadLot`,
    `sPrimerPaintBatchLot`, `sFinalPaintBAtchLot`, `sFrontImpactPadLot`,
    `sCenterImpactPadLot`, `sHelmetCoverLot`, `sShroudLot`, `sTrimLot`,
    `sRailLeftLot`, `sRailRightLot`, `sFitBandLot`, `sApex2ComfortFrontLot`,
    `sApex2ComfortCrownLot`, `sApex2ComfortSideLot`, `sApex2ComfortRearLot`,
    `sApex2ComfortNapeLot`, `sAccessoryStrapLot`, `sAnchorPDxTBondedLot`,
    `sTrimAdhesiveLot`, `sTrimSealantAdhesiveLot`, `sSleeveNapeLot`,
    `sAccessoryStrapRightLot`, `sHelmetBagLot`.

> The traceability tab name embeds the lot id (`LOT LG-01 Traceability`); parser
> locates the traceability sheet by pattern, not exact name.
> **[TO INVESTIGATE — B6]** Whether paint/primer/trim lots have documentation in a
> *different format* (may be excluded or handled specially in the document phase).

## 5. Function 1 — Database sanity check

**Databases** live on the network, opened **strictly read-only**.

**Product → database applicability**
- The set of applicable DBs depends on the product. **Bonding DB is conditional:**
  if the file has **no** bonding lot at all, skip the Bonding DB; if **some** rows
  have it and others don't, the file is **incomplete → flag**.
- **[TO INVESTIGATE — A4]** exact rule that maps Product/Part number → which DBs
  (Molding/Assembly/Quality) apply. Likely a small config table keyed by part number.

**Join keys** (per the discussion): each DB is matched to a file row by either
**SerialNumber** or **Mold-ID**, depending on the database.

**Field-to-DB mapping (draft, to confirm — A2)**

| File field | Likely DB | Join key |
|---|---|---|
| Mold-ID, molded `Weight`, molding date | Molding DB | Mold-ID |
| SerialNumber, `iCombatWeight`, component lots, `Item`, `Size`, `dtAssembled` | Assembly DB | SerialNumber |
| bonded lots (`sAnchorPDxTBondedLot`, adhesives) | Bonding DB | SerialNumber / Mold-ID |
| inspection result, `dtInspected` | Quality DB | SerialNumber |

> Not every file field lives in every DB — each field is only checked where it
> actually exists. Some dates come *from* the DB and are compared back to the file.

**Checks performed**
1. **Existence** — each serial/mold-id in the file exists in the applicable DBs.
2. **Field equality** — for fields present in a DB, file value == DB value.
   - Text compared **trimmed + case-insensitive** (G21).
   - Dates compared as dates (ignore time-of-day unless told otherwise).
3. **Weights vs Base information** — molded `Weight` within
   [`Min molded Weight`, `Max molded Weight`]; `iCombatWeight` within
   [`Min assembled Weight`, `Max assembled weight`]. Comparison is **within the
   tolerance defined by the Base information tab** (G19). Part number (`Item`)
   must equal `Product Part number`.
4. **No duplicates** — no duplicate `Mold-ID` and no duplicate `SerialNumber`
   **within the file** (DBs are assumed duplicate-free) (G20).
5. **Tab cross-check** *(recommended — H23)* — every serial in "Serial Number List"
   has a Traceability row and vice-versa; counts match; no extras.

Each check yields a per-serial, per-field verdict: **Correct / Incorrect /
Unknown** feeding the report.

## 6. Function 2 — Hand-over package + document validation

### 6.1 Bill of Material from Arena PLM (REST API)
- Base URL `https://api.arenasolutions.com/v1`. **Login** `POST /login`
  (email, password, optional `workspaceId`) → returns `arena_session_id`; send it
  as a header on every subsequent call.
- **Resolve part number → item:** `GET /items?number=<Product Part number>` → `guid`.
- **BOM:** `GET /items/<guid>/bom` returns a **single-level** list of lines, each
  with `item.number` and `item.guid`. To build the **full** BOM, **recurse** into
  each child guid. Cache results; mind the **24h request rate limit**.
- Output: a flat set of all BOM part numbers (all levels) used for cert matching.
- **[TO INVESTIGATE — F18]** how deep to recurse (full explosion vs N levels);
  credentials/workspace to use (dedicated integration user recommended).

### 6.2 Locating documents in the supplier tree
- Root: `M:\Armor\Newport\QUALITY\Incoming_Inspection\Supplier\<supplier>\...`.
  The supplier is **not** in the file, so the app **searches the whole tree**.
- **A lot's PDF is identified by the lot/PO string being part of the file name**
  (E14). POs are **unique per supplier** so no cross-supplier collision (E16).
- Lot values carry **PO prefixes `PR`, `PE`, or `PC`** and may contain **two lots**
  separated by `/` (B8) — each handled independently.
- **Date handling (E15):** the date is *sometimes* in the file name.
  When the file name alone is ambiguous:
  1. pick the candidate whose **file-creation date is closest** to the lot's date,
  2. open it and OCR-scan for a **handwritten inspection date** to confirm,
  3. if still unresolved, take the **most probable** file and **flag it Unknown /
     needs human validation**.

### 6.3 Building the hand-over package
- **One folder per lot column/category** (e.g. all Harness docs in one folder, all
  paint docs in another), each containing **all documentation for that lot** (B9).
- The **final report sits at the root** alongside the category folders.
- Per lot there is normally **one PDF containing both the COC and the PO** (C10).
  A lot spans ~500–1000 serials; in practice **1–2 (max ~5) lots per category**.

### 6.4 Validating each document
For every gathered PDF (OCR the scanned pages first):
1. **Contains both** a Purchase Order and a Certificate of Conformity.
2. **PO match** — when the COC lists a PO, it must equal the **PO number embedded
   in that lot's value**. Allow **leading zeros between the letter prefix and the
   digits** (C11) — e.g. `PR018950` matches `PR18950`.
3. **Part-number vs BOM** — at least one part number found on the document must be
   in the BOM.
4. **Partial part numbers** — if a COC part number is a **partial (6 chars: 5
   digits + a dash)**, that prefix must match the **start** of a BOM part number
   exactly.
- Each document yields verdicts (present/missing, PO ok/mismatch, part-in-BOM
  yes/no, unresolved→Unknown) rolled up per lot and per serial.

## 7. Lot-string parsing (shared, high-risk)

Real values are inconsistent: `PR18950 1-20-26`, `62117169`, `LM3245AE`,
`R06285`, `601191/512011`, `PR19082`, `PR1860110/22/25` (no space),
`TR5001 1-24-25`. Strategy:

- Split on `/` into **multiple lots** (but be careful: `/` also appears inside
  dates like `12/9/25` — only split when the right side isn't a date fragment).
- Extract an optional **PO token** matching `^(PR|PE|PC)0*\d+` (tolerating a
  missing space and leading zeros).
- Extract an optional trailing **date** in the many observed formats
  (`1-20-26`, `12/9/25`, `11-19`, `10/22/25`).
- Keep the **full original string** as the canonical lot id for file-name search
  (B8); use the PO token + date as *hints* for disambiguation.
- Unparseable/ambiguous values are surfaced as **Unknown** rather than guessed.

A dedicated unit-test corpus of real lot strings will guard this parser.

## 8. Function 3 — Report

Excel workbook written to the package root (**recommended — H22**; a PDF snapshot
can be added for the archive):
- **Summary sheet** — one line per serial number listing its issues (the
  confirmed-wrong and unknown items), for a fast triage read.
- **Detail sheet** — mirrors the lot file's layout, `SerialNumber` and `Mold-ID`
  kept intact; every other cell shows **Correct (green)**, **Unknown (yellow)**,
  or **Incorrect (red)** with matching cell fill.

## 9. Progress / status UX

Long operations must never look hung: a persistent **status bar** shows the
current stage and item (e.g. "Checking Assembly DB — serial 128/500",
"Searching supplier tree for PR18950", "OCR: harness COC page 2/3") with regular
updates from the worker thread. A cancel button stops cleanly.

## 10. Phased delivery

- **Phase 1 — DB sanity check + report.** Well-defined, high value, low external
  risk. Deliver: file parser, lot parser, DB adapters + product→DB mapping,
  all §5 checks, and the §8 report. (No network docs, no OCR.)
- **Phase 2 — Hand-over package + document validation.** Arena API integration,
  supplier-tree search, OCR, PO/part-number/partial matching, package assembly,
  and merge of those verdicts into the same report.

## 11. Open questions / to investigate

- **A4** Exact Product/Part-number → applicable-DB mapping (Molding/Assembly/Quality).
- **A2** Confirm the field↔DB↔join-key matrix in §5.
- **B6** Do paint / primer / trim lots have documentation (different format)?
- **F17/F18** Arena credentials + workspace; full-explosion depth for the BOM.
- **G19** The tolerance value(s) — confirm they are read from the Base information
  tab and how they apply to DB-vs-file weight comparisons.
- Access DB file names/paths and driver bitness on the target PC.
- Report format final call: Excel only vs Excel + PDF (H22).
- **Config & secrets** — which credential-storage option IT will accept (DPAPI
  vs Windows Credential Manager vs plain text), and whether DB/Arena access uses
  shared service credentials or the logged-in Windows user (see §3.1).

## 12. Risks

- **Messy lot strings** — the parser is the single biggest correctness risk;
  mitigated by a real-string test corpus and "flag as Unknown when unsure".
- **OCR reliability** on scanned certs — mitigate with preprocessing and by
  flagging low-confidence reads for manual check rather than asserting a result.
- **Arena rate limits** — cache BOM lookups; one BOM per part number per run.
- **Access driver bitness mismatch** — standardize on 64-bit engine + `.exe`.
