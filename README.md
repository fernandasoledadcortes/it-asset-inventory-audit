# IT Asset Inventory — Data Quality Audit

Auditing a messy ~1,000-row IT asset register before a hardware refresh and an
audit, and telling the IT Asset Manager **which data can be trusted and which
cannot** — delivered as a single, self-contained Excel workbook.

> Originally built as a take-home data challenge for a Report Analyst / BI role.
> The source dataset is a **fictional** asset register and is **not redistributed**
> here (see [Data](#data)). This repo shows the **code, method and findings**.

---

## The problem

An IT team is about to renew hardware and go through an audit. Their asset
register — one row per IT asset (laptops, desktops, phones, monitors, printers,
servers, virtual machines, network gear) — has grown messy: mixed formats,
duplicates, contradictions, values entered by different people over time. Before
anyone spends money on the refresh or trusts the numbers in the audit, someone
has to answer a simple question: **how reliable is each part of this data, and
why?**

The deliverable had to be **one Excel file** a non-technical IT Asset Manager
could open and act on — not a notebook.

## Approach

The guiding rule: **never silently fix or drop anything — flag it, record it,
and keep it reviewable.** The raw file is treated as read-only; every change is
logged cell by cell with the rule that justifies it, and the original is kept
alongside the clean version.

The pipeline runs in three numbered stages:

1. **Describe (EDA).** Define the grain (1 row = 1 IT asset), map the 34 columns
   (drop empty/duplicate ones, split the two different `Status` columns, name the
   unnamed review column), detect data types, mixed date formats, `mojibake`
   (`MÃ©xico` → `México`), unstandardised catalogues, impossible ranges, and
   measure completeness **only where a field actually applies** (a monitor with
   no OS is expected; a laptop with no OS is not).
2. **Clean.** Homogenise *everything* first (so `INUSE` and `In Use` count as the
   same value) and *only then* look for duplicates. Standardise identifiers
   (asset tags, hostnames, serials), dates, costs (converted to a common currency
   with a fixed, documented FX table), and lifecycle states.
3. **Contradictions & reliability.** Run explicit `chk_*` checks (one rule per
   check, with counts), then score reliability **per topic** (identification,
   costs, dates) rather than globally — so filtering out "to review" rows doesn't
   silently bias the sample toward, say, dropping every server.

Every assumption and decision is dated and justified in
[`docs/decisiones_y_supuestos.md`](docs/decisiones_y_supuestos.md).

## What the audit found

On the ~1,000-row fixture, after cleaning down to **985 distinct assets**:

- **Governance gap in naming.** Four different hostname conventions coexist but
  only one is documented → flagged as a data-governance finding, not a row error.
- **Devices that stopped reporting.** ~150 assets marked *In Use* hadn't reported
  to the management tools in 90+ days; several decommissioned/lost devices were
  **still reporting** (incl. laptops marked Lost/Stolen) — a security/recovery flag.
- **Stale assignments.** ~15 devices still assigned to terminated employees;
  reassigned devices whose email no longer matches the registered person.
- **Cost data you can't compare yet.** Costs in several currencies, some with the
  currency mislabelled (local amount tagged as USD) and ~19 atypical amounts
  (likely a lost decimal or wrong currency) — removed from statistics and flagged,
  not deleted.
- **The biggest real gap:** encryption status **unknown** for ~29% of end-user
  devices — which for an audit cannot be assumed either way.

The Excel deliverable turns this into a Cover + Executive Summary + Dashboard +
Analysis (native formulas) + Data Quality + Validation Rules + editable
Assumptions + Clean Data + Change Log + Original Data.

## Repo structure

```
codigo/
  00_comun/     config.py (paths, cut-off date, FX) + campos.py (field dictionary)
  01_describir/ EDA + first exclusions, with a change log
  02_limpiar/   contradiction checks (chk_* columns) on the clean base
  03_entregar/  builds the final English Excel (tabs, native formulas, charts)
docs/
  decisiones_y_supuestos.md   every assumption/decision, dated and justified
```

## How to run

```bash
python3 codigo/01_describir/explorar_activos_ti.py
python3 codigo/02_limpiar/revisar_contradicciones.py
python3 codigo/03_entregar/crear_excel_entrega.py
```

Run from the repo root, in order. Point `config.ARCHIVO_CRUDO` at your own asset
register (`crudo/DataChallenge.xlsx` by convention). Requires `pandas`,
`openpyxl`, `numpy`. Excel formulas are computed when the workbook is opened.

## Tools

Python (pandas, openpyxl, numpy) · reproducible, staged pipeline · cell-level
change log · data-quality / reliability framework · Excel reporting with native
formulas.

## Data

The source dataset is a **fictional** asset register (synthetic company, fake
domain) used for a take-home exercise. It is **not included** in this repo, so
the exercise material isn't redistributed. The code is written to run against any
asset register in the same shape.

## A note on how this was built

I used **Claude (Claude Code) as a programming assistant** — to scaffold the
project, write and run the code step by step, propose checks and options at each
decision, and draft the English texts. **The rules and decisions were mine**
(see the dated reasoning in `docs/`). This mirrors how I work day to day: AI for
speed and coverage, human judgement for every call that matters.
