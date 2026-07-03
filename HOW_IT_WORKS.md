# How It Works — Code Walkthrough

> A plain-language guide to the Helmet Lot Validation app, kept **in sync with the
> code**. Updated in the same commit as any code it describes. If you're reviewing
> the logic, read this first, then dip into the files it points to.
>
> Status: **scaffolding not started yet** — this document is populated as each
> module is built (see `DESIGN.md` §14 for the convention).

## General flow (end to end)

*(To be filled in as modules land. Intended narrative:)*

1. **Start & config** — the app loads the external config, resolves DB/file
   locations and credentials, and shows the main window.
2. **Pick file** — operator selects the lot workbook; it is parsed into typed rows.
3. **Phase 1 — DB sanity check** — the applicable databases are resolved for the
   product; each serial/mold-id is checked for existence, field equality, weight
   tolerances, part number, and duplicates. Verdicts are recorded per serial/field.
4. **Phase 2 — package & documents** — the full BOM is fetched from Arena; each
   sub-component lot's documents are located in the supplier tree, OCR-scanned,
   and validated (PO match, part-in-BOM, partial-match); files are copied into the
   hand-over package.
5. **Report** — a colour-coded Excel report (summary + detail) is written to the
   package root. Status bar reports progress throughout.

## Modules

*(One subsection per module, added as built. Template:)*

<!--
### `<module>/`
**Responsibility:** one or two sentences.
**Key functions/subs:**
- `function_name(args) -> return` — one-line purpose (inputs → action → output).
-->

## Key decisions & tricky logic

*(Documented as implemented. Expected entries:)*
- **Lot-string parsing** — how messy lot values are split/normalized.
- **Document date disambiguation** — how the right scan is chosen when the file
  name is ambiguous.
- **Partial part-number matching** — the 6-char (5 digits + dash) prefix rule.
- **Product → database applicability** — how the applicable DB set is chosen and
  how a partially-present bonding lot is flagged.
