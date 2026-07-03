# Dev Machine Setup Checklist

Prep for building the Helmet Lot Validation app in VS Code with the Claude Code
extension, on a machine that has real access so we validate against live data.
Work top-to-bottom; **Phase 1** items are needed first, **Phase 2** items can
follow later.

> Guiding rule (see `DESIGN.md` §13): keep sensitive bulk data on the PC. The
> assistant works from schema + a few sample rows + aggregate results, not full
> datasets. Nothing sensitive gets committed (`.gitignore` already blocks it).

## 0. The machine itself
- [ ] **Windows** PC (the app targets Windows — build on Windows to avoid surprises).
- [ ] You have permission to install software (or IT will install the items below).
- [ ] Decide **64-bit** as the standard (all tooling below must match — see §2 note).

## 1. VS Code + assistant
- [ ] Install **VS Code**.
- [ ] Install the **Claude Code** extension and sign in.
- [ ] Confirm the account's **data-retention / DPA terms** are acceptable for the
      data we'll touch (regulated data → verify before pointing tools at it).
- [ ] Clone this repo and check out the branch `claude/batch-config-scanner-app-mc0yq1`.

## 2. Python + build tooling  *(Phase 1)*
- [ ] Install **Python 3.12 (64-bit)**, "Add to PATH" checked.
- [ ] Confirm in a terminal: `python --version` and `pip --version`.
- [ ] (I'll create the project venv and `requirements.txt` during scaffolding.)

> **Bitness note (important):** Python, the Access ODBC driver (§3), and later the
> packaged `.exe` must **all be 64-bit** (or all 32-bit). Mixing them is the most
> common "can't connect to the DB" failure. We standardize on **64-bit**.

## 3. Microsoft Access connectivity  *(Phase 1)*
- [ ] Install **Microsoft Access Database Engine 2016 Redistributable (64-bit)**
      (needed even if Office isn't installed; provides the ODBC/OLEDB driver).
- [ ] Confirm the driver appears in **ODBC Data Sources (64-bit) → Drivers**
      ("Microsoft Access Driver (*.mdb, *.accdb)").
- [ ] Read access to each DB share; note the **UNC paths** and whether any DB has
      a **password**. (Goes into the config, encrypted.)
- [ ] Confirm access is **read-only** for our credentials.

## 4. Access & network reachability  *(Phase 1 + 2)*
- [ ] The four Access DB locations are reachable from this PC (Phase 1).
- [ ] The supplier tree root is reachable: `M:\Armor\Newport\QUALITY\Incoming_Inspection\Supplier` (Phase 2).
- [ ] The hand-over **output** location is writable (Phase 2).
- [ ] Outbound HTTPS to `https://api.arenasolutions.com` is allowed (Phase 2).

## 5. Sample data for validation
- [ ] **1–3 lot files** (`.xlsx`) — ideally anonymized like the one already shared.
      Include at least one that exercises edge cases (dual `/` lots, no bonding
      lots, odd PO prefixes `PE`/`PC`).  *(Phase 1)*
- [ ] A **read-only copy or pointer** to the relevant Access DBs, or agreement that
      I run read-only queries against the live ones.  *(Phase 1)*
- [ ] **3–5 real Certificate/PO PDFs** (scanned) covering good and messy cases —
      the OCR tuning depends entirely on these.  *(Phase 2)*
- [ ] Keep all of the above **out of git** (already covered by `.gitignore`; store
      under a local `samples/` folder).

## 6. OCR toolchain  *(Phase 2 — can defer)*
- [ ] Install **Tesseract OCR (64-bit)**; note the `tesseract.exe` path for config.
- [ ] Install **Poppler for Windows** (for PDF→image); note the `bin` path.
- [ ] Confirm both run from a terminal.

## 7. Arena PLM API access  *(Phase 2 — can defer)*
- [ ] A **login** for the API — ideally a **dedicated integration user**, not a
      personal SSO account.
- [ ] The **workspace ID** (or confirm default workspace is correct).
- [ ] Awareness of the **24h request rate limit** so test runs cache, not hammer.
- [ ] One known **part number** whose BOM we can fetch to validate end to end.

## 8. Credentials & secrets (security)
- [ ] Decide the secret-storage approach IT will accept: **Windows DPAPI-encrypted**
      (preferred), **Windows Credential Manager**, or plain config (last resort).
- [ ] Gather (but don't paste into git): DB passwords (if any), Arena login. These
      go through the Settings screen / encrypted config once built.

## 9. Ready-to-start gate
**Phase 1 can begin when:** §0–§3 done, DBs reachable read-only (§4), and at least
one sample lot file + DB query access exist (§5).
**Phase 2 adds:** §6, §7, the supplier tree + output paths (§4), and sample certs (§5).

---

### Quick "what I need from you to start Phase 1"
1. VS Code + Claude Code extension signed in, repo cloned to the branch.
2. Python 3.12 (64-bit) + Access Database Engine (64-bit) installed.
3. Read-only reachability to the Access DBs + their UNC paths / any passwords.
4. One or more sample lot files, and either DB copies or the OK to run read-only queries.
