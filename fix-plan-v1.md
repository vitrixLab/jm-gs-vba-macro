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
