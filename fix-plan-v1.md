# GlobalSmile v8.0 VBA Surgical Fix Plan v1

## Objective

Stabilize the v8.0 VBA project using the embedded workbook code as the source of
truth, while preserving the fail-closed accounting behavior. The plan separates
technical defects from accounting classification decisions and does not enable
production GL writes until the validation gates pass.

## Scope and evidence boundary

### Authoritative workbook

Use:

- `Global-Smile_2026-v8.0.xlsm`

The existing `_vba_extract/bas/` directory is not authoritative until its
provenance is verified. Its splitter currently points to the older
`Global-Smile_2026-v7.8_vba.txt`.

### Existing v8 source artifacts

- `modGLGate.bas`
- `modGLAggregation.bas`
- `modGLWorkbookMap.bas`

These files must be compared with the embedded VBA project before any patch is
applied to the workbook.

## Current status

- CDJ, CRJ, and GJ are the intended authoritative GL posting sources.
- PJ, PJ Non-Vat, and SJ are feeder journals and must not be posted twice.
- Combined source reconciliation is documented as passing:
  - Debit: `256,145.75`
  - Credit: `256,145.75`
  - Difference: `0.00`
- The expected calculation shape is `46 accounts x 12 months = 552 records`.
- COA coverage remains on HOLD for five GJ labels totaling `6,154.00`.
- Runtime execution in Excel is not yet certified.

## Fix sequence

### F1 — Establish extraction provenance

**Target:** extraction tooling and generated source directory.

1. Extract the VBA project directly from `Global-Smile_2026-v8.0.xlsm`.
2. Write the output to a versioned directory such as `extract_v80/`.
3. Split the extracted project into `.bas` and `.cls` files.
4. Record:
   - workbook SHA-256;
   - VBA project SHA-256, where available;
   - module names;
   - module line counts;
   - extraction warnings or failures.
5. Do not use the existing `_vba_extract/bas/` files as v8.0 evidence until they
   match the new extraction.

**Acceptance checks**

- Every embedded standard module, class module, worksheet module, and
  `ThisWorkbook` module is accounted for.
- The extracted v8 modules can be traced back to the authoritative workbook.
- No workbook binary is modified during extraction.

### F2 — Reconcile the embedded v8 modules with source artifacts

**Target:** extracted v8 modules and the standalone `modGL*.bas` files.

Compare the embedded modules with:

- `modGLGate.bas`
- `modGLAggregation.bas`
- `modGLWorkbookMap.bas`
- `modEngine.bas`
- `modPJAutomation.bas`
- `Module1.bas`
- `ThisWorkbook.cls`
- Graphify modules

Create a difference list before editing. Classify each difference as:

- expected version difference;
- stale source artifact;
- workbook-only change;
- actionable defect.

**Acceptance checks**

- The planned patch is tied to exact extracted module names and line locations.
- No source file is overwritten merely to make the diff look clean.

### F3 — Fix the CRJ writer/layout conflict

**Target:** `Module1.CreateCRJFromSales` and the central CRJ writer in
`modEngine`.

The current `CreateCRJFromSales` implementation writes fields to columns
`A:H`, while the v8 engine defines CRJ fields in:

| Field | v8 column |
|---|---|
| Date | D |
| Customer | F |
| Invoice | G |
| Cash | H |
| Output VAT | J |
| Exempt | K |
| Sales | L |
| Note | O |
| Reference | P |
| Sequence | Q |

**Required change**

Use one canonical writer. Prefer delegating `CreateCRJFromSales` to the
centralized `modEngine` CRJ synchronization routine. Do not maintain a second
column layout in `Module1`.

If delegation is not possible, rewrite the procedure to use the exact v8 CRJ
columns and populate the reference/sequence fields consistently.

**Acceptance checks**

- A test SJ row creates one CRJ row in the v8 layout.
- Date, customer, invoice, cash, VAT, sales, note, reference, and sequence are
  written to the columns consumed by the v8 aggregation engine.
- Re-running the same source row does not create a duplicate.
- The generated CRJ row is included in the expected month and balances.

### F4 — Resolve CDJ column Q and S handling

**Target:** `modGLAggregation.Matrix` and `modGLWorkbookMap.V8_CDJMap`.

The aggregation loop scans CDJ columns `F:S`, but the current map explicitly
handles `F:P` and `R` only. Columns `Q` and `S` therefore become ambiguous
activity whenever non-zero.

**Required investigation**

1. Inspect CDJ row 13/14 headers for columns Q and S.
2. Determine whether each is:
   - an account posting column;
   - a sundry/category column;
   - metadata that must not be aggregated.
3. For an account posting column, add an explicit COA mapping.
4. For metadata, exclude it from the posting loop rather than treating it as an
   unmapped account.
5. Preserve the fail-closed behavior for any genuinely unknown posting column.

**Acceptance checks**

- Every CDJ posting column is explicitly classified.
- Legitimate Q/S activity maps to a documented COA account.
- Metadata in Q/S cannot inflate debit, credit, or unmapped totals.
- The CDJ reconciliation remains `4,500.00 debit = 4,500.00 credit`.

### F5 — Make calculation-sheet generation fail closed

**Target:** `modGLAggregation.BuildV8CalcSheet`.

`BuildV8CalcSheet` can currently be called directly and build a 552-row sheet
without independently rejecting unmapped or invalid activity.

**Required change**

Before writing or clearing `GL_V8_CALC`, run the same validation controls used by
the gate:

- debit/credit difference within `TOLERANCE`;
- unmapped amount within `TOLERANCE`;
- invalid numeric/date count equals zero;
- matrix shape equals `46 x 12`.

On failure:

- do not clear an existing valid calculation sheet;
- do not write partial output;
- write a clear HOLD record to `GL_AUDIT`;
- return `False` and show the repository-standard user notification.

**Acceptance checks**

- Invalid or unmapped data leaves the prior calculation output intact.
- Valid data produces exactly 552 account-month rows.
- Direct calls to `BuildV8CalcSheet` cannot bypass validation.

### F6 — Harden initialization and required-sheet errors

**Target:** `modEngine`, `modPJAutomation`, and `ThisWorkbook`.

Replace broad `On Error Resume Next` around required workbook setup with explicit
checks for:

- `CDJ`;
- `CRJ`;
- `GJ`;
- `GL`;
- `SUPPLIERS DATA`;
- required source sheets used by the selected operation.

Initialization must fail visibly if a required sheet, dictionary, header, or
sequence source is unavailable.

**Required decision**

Choose one startup model and document it:

1. initialize both `modEngine` and `modPJAutomation` in `Workbook_Open`; or
2. keep PJ automation lazy, but expose and log its initialization state.

**Acceptance checks**

- Missing required sheets produce a clear error and a non-ready state.
- A failed cache load cannot be mistaken for successful initialization.
- Workbook open behavior remains usable when optional sheets are absent.

### F7 — Keep the five COA classifications on accounting HOLD

**Target:** `V8_GJMap` and workbook COA, only after accounting approval.

Do not add speculative aliases for:

- Medical Equipment — `2,992.50` credit
- Cost of Revenue — `1,139.00` debit
- Supplies — `2,000.00` debit
- Bank Charge — `15.00` debit
- Charges — `7.50` debit

These require an explicit accounting decision. Once approved, add documented
aliases or update the 46-account COA and record the mapping decision in the
review evidence.

**Acceptance checks**

- Before approval, these values remain visible in `GL_AUDIT` and validation
  remains HOLD.
- After approval, unmapped activity reaches zero and the resulting account
  classifications are traceable.

### F8 — Runtime validation in Excel

Run the corrected embedded VBA project in Excel against a copy of the
authoritative workbook.

Execute and record:

1. `Workbook_Open` initialization.
2. `ValidateV8Structure`.
3. `ValidateGLConsistency(2026)`.
4. `BuildV8CalcSheet(2026)` through the validated path.
5. `RefreshGL(2026)` while the five mappings remain unresolved.
6. The CRJ writer test from an SJ row.
7. CDJ Q/S posting tests.

**Acceptance checks**

- HOLD prevents production `GL` overwrite.
- `GL_AUDIT` records source totals, unmapped amount, invalid count, and status.
- No duplicate feeder-journal posting occurs.
- Validated output has the expected 552 records.
- Runtime results agree with independently calculated journal totals.

## Regression matrix

| Case | Expected result |
|---|---|
| Missing CDJ/CRJ/GJ/GL sheet | Clear HOLD/error; no write |
| Missing required header | Clear HOLD/error; no write |
| Invalid numeric value | Invalid count increases; no GL write |
| Invalid CRJ date/log date | Invalid count increases; no GL write |
| Unmapped GJ account | Listed in `GL_AUDIT`; HOLD |
| Non-zero CDJ Q/S metadata | Ignored if metadata, not posted |
| Non-zero CDJ Q/S account | Explicitly mapped and included |
| CRJ generated from SJ | Written using canonical v8 layout |
| Duplicate source row | No duplicate CRJ entry |
| Feeder journals included | Not posted independently |
| Debit/credit mismatch | HOLD; existing GL preserved |
| Valid 46-account population | Exactly 552 calculation records |
| Graphify execution | Does not alter accounting calculations |

## Definition of done

The v8.0 fix is complete only when:

- extraction provenance is verified for the authoritative workbook;
- CRJ writing and CRJ reading use the same layout;
- CDJ Q/S columns are explicitly classified;
- calculation-sheet generation is fail closed;
- required initialization failures are visible;
- runtime Excel tests pass;
- all unmapped activity is either explicitly approved and mapped or remains
  intentionally on HOLD;
- production `GL` writing is enabled only after all required gates pass.

## Non-goals

- Do not modify `Graphify` to perform accounting aggregation.
- Do not silently map unresolved accounts to a convenient existing account.
- Do not overwrite the production `GL` sheet while validation is HOLD.
- Do not treat a balanced journal population as proof of complete COA coverage.

---

# Part B — Graphify Per-Webpage Modularization + Fault-Finding Observation + Analytics Log

> Scope: `graphify.html`, `graphify_vba.html`, `Graphify.bas` in THIS repo (`globalsmile`,
> branch `main` @ `1da1534`, remote `vitrixLab/jm-gs-vba-macro`).
> Part A (above) stays authoritative for the v8.0 VBA surgical fix. This Part B only
> modularizes the Graphify presentation/demo layer. It must not alter accounting
> behavior (`modGLGate` / `modGLAggregation` / `modGLWorkbookMap` untouched, G7 holds:
> Graphify stays downstream of accounting).

## B0. What was reviewed (evidence)

- `Graphify.bas` (229 lines, `modGraphify`): `Sub Graphify()` + 3 helpers
  (`PickNumericColumn`, `FirstNumericRow`, `ColumnLetter`) + `Fail:` handler that only
  `MsgBox`es. Resolution order: (1) multi-cell selection wins, (2) `CHART_COLUMN = "I"`
  (>= 2 numerics), (3) auto busiest-numeric-column, (4) `MsgBox "No numeric data found"`.
  Chart `GraphifyChart` reused (series rebuilt), docked at `E2`, categories `R1..Rn`,
  live-bound (`srs.Values = dataRng`, blanks stay gaps). No logging/analytics.
- `graphify.html` (63 lines): single static Chart.js v4 page, hardcoded
  `sampleValues = [23,45,12,78,34,56,29,80,15,62]`, labels `R1..Rn`. No header-nav,
  side-nav, error/empty state, fault log, or analytics.
- `graphify_vba.html` (97 lines): static highlight.js visual of the OLD broken v7.8
  listing (`ws.Charts.Add`, `Shapes("GraphifyChart.chart.8")`, missing `End Sub`).
  Documents the bug; is NOT the fixed `Graphify.bas`. No nav, no analytics.
- Context: `PLAN v1.md` mandates Graphify stays downstream
  (`modGLAggregation -> modGraphify -> GL Charts`). `status_v2.md` keeps G7 UNVERIFIED;
  Part A regression requires "Graphify execution does not alter accounting calculations".

## B1. Target: per-webpage component split

Keep two static pages. Split each into small includes (header-nav, side-nav, chart,
code-visual, observability = one file each):

```
graphify.html                      # shell only (~40 lines): head + includes + boot
partials/
  header-nav.html                  # NEW: brand, links (Demo / VBA Code / GL Charts), version badge v7.9-fixed
  side-nav.html                    # NEW: Chart Demo / VBA Source / Mapping (PLAN v1) / Gates (status_v2)
  chart-panel.html                 # EXTRACT from graphify.html: card + <canvas> + dataset switcher
  footer.html                      # NEW: G7 notice (chart only) + workbook hash slot
assets/
  graphify-config.js               # NEW: datasets (sample + 46x12-shape demo + presets I/F/G)
  graphify-chart.js                # NEW: render/reuse mirroring Graphify.bas (selection->col I->auto->empty)
  graphify-analytics.js            # NEW: fault + analytics log (see B3)
  graphify.css                     # EXTRACT inline <style> from both pages
graphify_vba.html                  # shell only: header-nav + side-nav + code-visual + fault log
partials/
  code-visual.html                 # EXTRACT <pre><code>; tabs: v7.8-broken vs v7.9-fixed (Graphify.bas)
  fault-log.html                   # NEW: visible fault-finding panel (mirrors MsgBox paths)
```

Rules: no accounting math in `assets/*.js` (charting only); `CHART_COLUMN="I"` default
kept; categories stay `R1..Rn`; chart reused by id; both pages share nav/footer/CSS.

## B2. Fault-finding observation system (per webpage)

Mirror every `Graphify.bas` failure path as a visible, logged fault (today MsgBox-only):

| VBA path (`Graphify.bas`) | Web fault code | Panel behavior |
|---|---|---|
| Not a worksheet (`TypeOf ActiveSheet`) | `NOT_WORKSHEET` | fault-log row + empty chart + guidance |
| No numeric data (all 3 resolutions fail) | `NO_NUMERIC_DATA` | empty state ("Select a column and run again") |
| No numeric block (`firstDataRow = 0`) | `NO_NUMERIC_BLOCK` | empty state with resolved source (`srcDesc`) |
| Unexpected (`Fail:` Err.Number/Description) | `RENDER_FAILED` | error state with code + Retry button |
| Success | `RENDER_OK` | chart + source caption (column I / selection / auto) |

`partials/fault-log.html` = table (time, page, code, detail, source) fed by a
ring buffer (cap 200, `localStorage` persist). `?debug=1` expands detail rows.
Excel-side parity (optional `GraphifyAudit` sheet writer) needs Part A gate approval
first — do NOT touch any workbook binary in this Part B.

## B3. Log-every-analytics contract (static-site sized)

`assets/graphify-analytics.js` (vanilla, no dependency, <150 lines):

```js
logPageView(page)                             // on load, per page
logChartRender({page, source, nPts, code})    // RENDER_OK + srcDesc + point count
logFault({page, code, detail})                // every fault above; 100% sampled
logAction(action, meta)                       // dataset switch, tab switch, debug toggle
// transport: in-memory ring + localStorage + console;
// optional POST to /api/analytics/event when hosted
// payload: {t, page, code, source, detail} — counts only, no workbook values
```

Verify: load each page with `?debug=1`, force each fault (empty dataset, bad column,
render throw) → each appears in fault-log with page/code within 60s and survives reload.

## B4. Execution (3 small steps, no workbook touch)

- [ ] **B-Step 1 — extract, no behavior change:** create `partials/` + `assets/graphify.css`,
  move inline style/code blocks verbatim; both pages render byte-equivalent.
- [ ] **B-Step 2 — components:** add `header-nav`/`side-nav`/`footer` includes,
  `graphify-config.js` + `graphify-chart.js` (dataset switcher incl. 552-shape demo),
  `code-visual` tabs (v7.8-broken vs v7.9-fixed from `Graphify.bas`).
- [ ] **B-Step 3 — observability:** add `fault-log.html` + `graphify-analytics.js`
  (ring 200, localStorage, `?debug=1`); map all 5 fault codes; verify matrix below.
- Exit: each page a <60-line shell; shared nav identical; all faults log + persist;
  **no `.xlsm` modified** (`git status` shows only html/partials/assets); Part A gates OK.

## B5. Regression matrix (adds to Part A table)

| Case | Expected |
|---|---|
| Empty/non-numeric dataset | `NO_NUMERIC_DATA` in fault-log + empty state, no throw |
| Single numeric cell (< MIN_NUMERIC=2) | `NO_NUMERIC_DATA` + guidance |
| Render throw (bad ctx) | `RENDER_FAILED` with code + Retry, prior chart kept |
| Dataset switch | `logAction` + `RENDER_OK` with new source caption |
| Reload with `?debug=1` | fault history restored from localStorage |
| GL sheets / `.xlsm` | untouched; Part A accounting evidence still holds |
