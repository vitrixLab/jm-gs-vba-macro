# Gate 2 Evidence Report — Global-Smile_2026-v8.4-fix.xlsm

**Commit under review:** `703a3213ab926ed10415de48c4ab5940deb81554`
**Candidate:** `Global-Smile_2026-v8.4-fix.xlsm`
**Decision:** **HOLD — do not promote.** The candidate's journals still fail the
fail-closed gate (Dr 4,431,497.52 / Cr 4,972,059.43, D−C −540,561.91, unmapped
698,740.35, invalid 559), the GL is still the stale v8.3.3 snapshot, 7 sundry
labels require accounting classification decisions that were not supplied, and
two data-integrity defects were found (132 zeroed day cells in the PJ source
journal; one accidental Non-Vat row). No commit, push, merge or promotion was
performed. Nothing in this report was applied to any workbook.

Method: all figures below are produced by independent read-only Python
aggregation (`recon_gl.py`, `crosscheck.py`, `gate2_evidence.py`,
`gate2_resid.py`, temp dir) that does not use the workbook's VBA. Where the
independent path and the workbook's own recorded gate run agree, both numbers
are shown. Excel was not run: compile check, runtime gate re-run, second-refresh
determinism and the live 46×12 write-back are **NOT RUN** and are not claimed.

---

## 1. Preservation (directive item 1)

| File | SHA-256 | Bytes | Modified |
|---|---|---|---|
| `Global-Smile_2026-v8.4.xlsm` (production, must not change) | `9DF84B7B86482C9D8B7C7FCC6EF4842824E4BF8F47CDFEB848A1BC8D6D56B11C` | 2,930,824 | 2026-10-13 00:54 |
| `Global-Smile_2026-v8.4-fix.xlsm` (candidate) | `CEC21E9AE4E2BE0412D06DE8EAB90BE3C063F75BC4F98E9F44376416745DB57B` | 2,900,397 | 2026-10-14 11:30 |
| `Global-Smile_2026-v8.3.3.xlsm` (certified control) | `6754126327B59397818E5C58C5256CE96ADD760D5FE126F522F8B58CAFA467EE` | 1,239,325 | 2026-10-02 23:44 |

Production SHA is **unchanged** since baseline (PASS). Working tree was clean
at HEAD `703a3213` before this review; this report and
`cdj_deletion_ledger.csv` are the only new files.

## 2. Sheet structure (directive item 2)

Candidate has **17 sheets, names identical** to production
(`PJ, PJ Non-Vat, CDJ, SJ, CRJ, GJ, GL, GL_V8_CALC, SUPPLIERS DATA, CRJ2, PJ2,
PJ Non-Vat2, CDJ old, SJ2, COA_MAP, GL_AUDIT, V8.3_FINAL_REVIEW`). No sheet was
added, removed or renamed (PASS). The existing GL layout (46 blocks × 12
months, stride 13, first row 13) is untouched (PASS). `COA_MAP`, `GL_V8_CALC`
and `GL_AUDIT` remain as before (non-load-bearing; `GL_V8_CALC` still 552 rows
with Σ|Δ| vs GL = 0.0000).

## 3. Candidate VBA delta (smallest-change audit)

Four modules differ from the shipped baseline; all four match the committed
`gate2_work/vba_patched/` exports (byte differences are extractor encoding of
non-ASCII characters only). The other 22 modules are byte-identical to baseline.

| Module | Change | Verdict |
|---|---|---|
| `modEngine` | **Compile fix**: `ValidateGLConsistency`/`BuildV8CalcSheet` now defined (:1434,:1438) as wrappers over `Legacy_*` which self-assign correctly. CDJ reference columns moved 25/26 → 27/28 with legacy fallback; `FindCDJFreeRow` replaces the stale AA14 next-row pointer; CDJ month-marker helpers added. | **Keep** — fixes the C1 compile break with the smallest edit (modGLGate/modGLScanEngine untouched, still call the same names). Static evidence only: runtime compile NOT RUN. |
| `modGLWorkbookMap` | `V8_GJMap` gains 9 aliases: 5 visiting variants (incl. `VISITING CLINICIANS FEE`) → `Clinician's Fee-Visiting`; 4 corporators variants (incl. `CORPORATOS`, `CLINACIANS`) → `Clinician's Fee-Corporators`. | **Keep the 8 mechanical variants** (apostrophe/spacing/truncation only). `VISITING CLINICIANS FEE` and the un-mapped `PARTNERS CLINICIANS FEE` family are accounting classifications — see §6. |
| `modPJSync` | `SyncToCDJ_NonVat` disabled (`Exit Sub`, decision D3: `PJ Non-Vat2` mirrors PJ → double-count). | **Keep** — prevents double posting; lookups/totals preserved. |
| `Sheet9` (PJ Non-Vat2) | Undeclared constants declared (`Option Explicit`); reference generation disabled (`If False Then`). | **Keep** — fixes the S1 silent-Empty defect; same D3 rationale. |

No other VBA was touched. `modGLAggregation.bas` remains absent from the
workbook (as designed after the C1 fix).

## 4. CDJ row-level deletion ledger (directive item 4)

Full ledger: **`cdj_deletion_ledger.csv`** (159 signature groups, one row per
removed copy with surviving-twin rows, refs, and PJ source hits).

Production 671 data rows / 423 unique `(particulars, cash, sundry)` signatures →
candidate 505 data rows / 424 signatures. Removed-class rows total **171**;
added-class rows total **5**.

| Class | Rows | Evidence |
|---|---|---|
| Excess duplicate copies removed, **surviving twin present** | **167** | Every removed copy has ≥1 byte-identical twin (same particulars, cash, sundry) remaining in the candidate. Largest: `MDI GROUP HOLDINGS INC` ×4→×2 (−169,845.93 twice), `FACIAL CARE CENTRE` ×16→×8, `SMARTPETRO` ×2→×1, `VPHARMA` r15+r232 → r232 kept, etc. |
| Transactions **orphaned** (no twin anywhere) | **0** | Verified both by signature multiset and by reference: all 173 lost col-AA refs have a surviving twin row carrying the same transaction. The lost refs were per-copy sequence numbers minted by the sync, not per-transaction identities. |
| Old monthly TOTAL rows recomputed in place (blank-E, month label in col C) | 4 | `r111/r193/r408/r516` rewritten from the old partition sums to the new deduplicated sums (e.g. r111 −968,712.43 → −847,946.13 = Jan partition sum). Inert for the engine (`modGLScanEngine.bas:179` skips blank or `TOTAL` titles) — same convention as production. |
| New monthly TOTAL rows added | 4 | May/Jun partition totals (`r612` −889,584.24, `r674` −601,955.83) + the two rewritten above complete the 6-month set. |
| New transaction row added — **Non-Vat sync artifact** | **1** | `r15` = FEDERAL BRENT RETAIL INC −1,500.00, ref `AB=PNV-00002`, log `Inserted Non-Vat Row 8 | EWT: 0 | 10/14/2026 10:31`. Posted by the still-live `SyncPJNonVatToCDJ` during the candidate build (edit-event side effect). Row is balanced (cash −1,500 + N 1,500, nets 0) and traces to a genuine `PJ Non-Vat` r8 invoice (`PNVINV-20260802-195122-00002`), **but production CDJ has 0 PNV rows and the sibling invoice `PNV-00004` (also ₱1,500) was not posted** — see manifest item 9. |
| PJ source-journal cells modified | **132 (col C only, rows 8–183)** | **Unauthorized source change — REVERT.** All 132 cells are PJ day-of-month values zeroed (`'03'→0`, `20→0`, …). No GL benefit; corrupts the feeder journal. |

CDJ row-balance check: **0 rows** with col-X residual > 0.01 (production had 5
rows totalling +11,190.00; the candidate's repair also cleared those).

**Ledger conclusion:** the CDJ data repair is economically a pure duplicate
removal (167 excess copies, every one explained, plus 4 monthly TOTAL rows
recomputed in place), plus one accidental Non-Vat sync row (r15). The PJ col-C
zeroing is not explained by any reconciliation need and must be reverted before
any promotion.

## 5. CRJ 559 invalid-row investigation (directive item 5)

Candidate CRJ is **byte-identical** to production CRJ. Structure re-proven:

| Partition | Rows | Invoices | Σ col H | Col-D TOTAL row | Diff |
|---|---|---|---|---|---|
| JANUARY | 10–114 | 105 | 593,273.12 | r115 = 593,273.12 | **0.00** |
| FEBRUARY | 116–196 | 80 | 450,514.28 | r197 = 450,514.28 | **0.00** |
| MARCH | 198–316 | 118 | 837,607.99 | r317 = 837,607.99 | **0.00** |
| APRIL | 318–389 | 71 | 421,057.12 | r390 = 421,057.12 | **0.00** |
| MAY | 391–486 | 95 | 548,191.71 | r487 = 548,191.71 | **0.00** |
| JUNE | 488–581 | 93 | 606,743.55 | r582 = 606,743.55 | **0.00** |
| **Total** | | **562** | **3,457,387.77** | | **0.00** |

Findings:

1. **The 559 "invalid" rows are valid invoices with no usable date.** The engine
   (`modGLScanEngine.bas`, `ScanCRJ`) dates rows from **col O only**; col O
   carries SJ import *timestamps* (`… | 8/31/2026 1:50:49 PM`), not transaction
   dates. 559 of 562 invoice rows have no col-O value → `inv` → HOLD.
2. **The 3 dated rows are mis-monthed.** `r10/r11/r12` (invoices 1819/1820/1821,
   Σ H = 10,160.00) live in the **JANUARY** partition but their col-O timestamps
   are August 2026 → the engine posts them to **August**. This is a wrong-month
   posting, not a fixable data gap.
3. **The month source is provable and already present:** the six col-D TOTAL
   rows partition all 562 invoices exactly (each partition sums to its TOTAL to
   the cent; Σ = 3,457,387.77 = the workbook's own certified CRJ total). The
   legacy `GLScanCRJ` in `modEngine` already uses this partition method.

**Stop-and-report item (mechanics, evidence complete, remediation not
applied):** CRJ month attribution must switch from col-O timestamps to the
col-D partition month (either a ~15-line change in `ScanCRJ`, or a data repair
writing the partition month into col D of the 559 rows). Which one is acceptable
is a maintainer decision; both are specified, neither was executed.

## 6. Unmapped reconciliation — ₱698,740.35 (directive item 5, second half)

The prior report's candidate figure **reproduces exactly** once the candidate's
patched `V8_GJMap` aliases are applied: **698,740.35** across 107 row-groups /
7 labels. The gate's unmapped metric is Σ(|cash|+|allocations|) per unmapped
row; the GL-imbalance contribution is the row's *sundry residual*. Both are
given; the residual column sums to the engine's D−C **to the cent**:

| Label (CDJ col S) | Rows | Unmapped metric | Residual → D−C | GL account? |
|---|---|---|---|---|
| Representation | 93 | 438,291.68 | 353,461.04 | **No account exists in the COA** |
| Partners Clinicians Fee | 4 | 104,817.73 | 104,817.73 | No (classification: partners vs due-to-clinicians?) |
| SSS,HDMF AND PHIC ER SHARE | 5 | 119,828.96 | 47,180.80 | No exact title; `Government Contributions ER Share` exists |
| Partners Clinicians Fee- Corporators | 1 | 17,319.64 | 17,319.64 | No |
| Incentives | 1 | 10,667.52 | 10,667.52 | No |
| Outside Services | 2 | 4,200.00 | 4,200.00 | No |
| Training and Seminar | 1 | 3,614.82 | 2,915.18 | No |
| **Total** | **107** | **698,740.35** | **540,561.91** | |

Tie-out: engine D−C = −540,561.91 = −(353,461.04 + 104,817.73 + 47,180.80 +
17,319.64 + 10,667.52 + 4,200.00 + 2,915.18). ✓

The 9 mechanical clinician-fee variants (₱930,534.90 on production data) are
now mapped by the candidate's alias patch and no longer appear. What remains is
**accounting classification, not spelling**. Per the stop rule these are not
guessed: each needs an owner decision (map to an existing account, or add an
account to the COA — adding accounts is out of Gate 2 scope).

**Blocker B1:** classification of the 7 labels above (₱540,561.91 of imbalance;
₱698,740.35 of unmapped metric). Notably: Representation (₱353,461.04) has no
home in the current COA at all; SSS/HDMF/PHIC ER SHARE probably belongs to the
existing `Government Contributions ER Share`, but that is an accounting
assertion this review will not make on its own.

## 7. 46 accounts × 12 months reconciliation (directive item 7)

Independent aggregation of the candidate's posted journals vs the GL sheet:

| Metric | Production v8.4 | Candidate v8.4-fix | Certified control v8.3.3 |
|---|---|---|---|
| Engine-equivalent Dr | 4,423,330.72 | 4,431,497.52 (post-alias) | 258,671.00 |
| Engine-equivalent Cr | 6,108,350.66 | 4,972,059.43 | 258,671.00 |
| D−C | −1,685,019.94 | **−540,561.91** | 0.00 |
| Unmapped | 1,888,683.84 | **698,740.35** | 0.00 |
| Invalid | 559 | **559** | 0 |
| GL stored Dr/Cr | 258,671.00 / 258,671.00 | 258,671.00 / 258,671.00 | 258,671.00 / 258,671.00 |
| Accounts GL ≠ journals | 18 / 46 | **18 / 46** | **0 / 46** |
| Month cells differing | 90 / 552 | **88 / 552** | **0 / 552** |

The candidate reduced the imbalance by ₱1,144,458.03 (dedup + clinician-fee
aliases) but is nowhere near the ≤ ₱0.01 acceptance threshold. The GL was
correctly **not** overwritten: the workbook's own last gate run is recorded in
`GL_AUDIT` row 9 — `2026-10-14 11:24:41 HOLD | Debit=3,468,245.56 |
Credit=4,972,059.43 | Unmapped=1,661,992.31 | Invalid=559` — i.e. the run
happened *before* the alias patch was injected (the alias-patched engine, run
today, would report the §6 numbers). Fail-closed behaviour: **PASS** (HOLD
recorded, GL untouched). No post-alias runtime evidence exists: **NOT RUN**.

**Status of criterion 7: FAIL** (0/46 accounts agree; control v8.3.3 shows the
test itself is achievable).

## 8. Change manifest (directive item 6 — proposed, NOT applied)

Already in the candidate — keep:
1. `modEngine` compile fix + CDJ ref-column migration + `FindCDJFreeRow`.
2. `modGLWorkbookMap` — 9 clinician-fee aliases (8 mechanical variants; the
   `VISITING CLINICIANS FEE` alias is a classification call — see B1).
3. `modPJSync` D3 disable + `Sheet9` constant/ref-generation fix.
4. CDJ duplicate removal (167 excess copies + 4 recomputed monthly totals, fully ledgered in §4).
5. CDJ col-X row-balance repair (5 rows, ₱11,190.00).

Required before any promotion — not applied, needs approval:
6. **Revert PJ col-C zeroing** (132 cells, rows 8–183) — restore production values.
7. **Accounting decisions for the 7 labels** (§6, blocker B1). No mapping is
   proposed here because each choice changes the financial statements.
8. **CRJ month attribution** — partition-anchored (§5). Two equivalent
   implementations specified; neither executed (blocker B2: maintainer chooses
   VBA change vs data repair).
9. **Resolve the Non-Vat ingestion inconsistency** (blocker B4): candidate CDJ
   r15 posted `PNV-00002` (₱1,500) as an edit-event side effect while sibling
   invoice `PNV-00004` (₱1,500) was not posted and production never posted any
   PNV row. Decide: post both (complete the Non-Vat source) or remove r15
   (restore the "CDJ = production minus duplicates" invariant). Either is a
   two-cell change; the inconsistency itself must not ship.
10. After 6–9: rerun the gate → expect PASS with Dr = Cr, unmapped 0, invalid 0;
    then re-run the independent 46×12 reconciliation (must be 0/46 differing);
    then prove a second refresh is identical; then re-verify the production
    workbook SHA. All of step 10 is **NOT RUN** today.

## 9. Status table (directive item 9)

| # | Item | Status |
|---|---|---|
| 1 | Production workbook preserved (SHA) | **PASS** |
| 2 | No sheet add/remove/rename/redesign | **PASS** |
| 3 | GL layout kept as primary ledger | **PASS** |
| 4 | CDJ deletion ledger with references/source evidence | **PASS** (171/171 removed-class rows explained: 167 duplicates with twins + 4 recomputed totals; 0 orphans) |
| 5 | 559 CRJ rows + ₱698,740.35 unmapped investigated | **PASS** (investigated and reconciled exactly; remediation blocked, see B1/B2) |
| 6 | Fail-closed gate not bypassed | **PASS** (HOLD recorded; GL untouched) |
| 7 | Smallest possible VBA/workbook changes | **PARTIAL** (candidate edits are minimal and justified; but it also modified 132 PJ source cells without justification and left one accidental Non-Vat row — revert items 6 and 9) |
| 8 | Independent 46×12 reconciliation ≤ ₱0.01 | **FAIL** (0/46 agree; 18/46 differ) |
| 9 | Journals balance / GL matches journals | **FAIL** (D−C −540,561.91; GL stale) |
| 10 | Second refresh identical, original unchanged | **NOT RUN** (no Excel runtime in this review) |
| 11 | Runtime compile / live gate after alias patch | **NOT RUN** (static evidence only) |
| 12 | Certified control v8.3.3 still passes independently | **PASS** (0/46, 0/552 reproduced) |

**Overall: HOLD.** A passing component test (compile fix, alias table, dedup
ledger) is not a passing workbook: the candidate's GL does not equal its
journals and cannot until B1 and B2 are resolved.

## 10. Blockers (stop-and-report, directive item 10)

- **B1 — Accounting classification (7 labels, ₱698,740.35 / residual ₱540,561.91).**
  Representation has no COA account; Partners Clinicians Fee (variants);
  SSS,HDMF AND PHIC ER SHARE; Incentives; Outside Services; Training and Seminar.
  Each needs an owner decision (existing account, or a COA change outside Gate 2 scope).
- **B2 — CRJ month attribution choice.** Partition-anchored attribution is proven
  exact (§5); the fix must be chosen (engine change in `ScanCRJ` vs data repair
  of 559 rows) and then executed under change control.
- **B3 — Unauthorized PJ source edit.** The 132 zeroed day cells must be reverted
  (item 6); their origin is not documented in any commit message or audit row.
- **B4 — Partial Non-Vat ingestion.** Candidate CDJ r15 posted `PNV-00002` as a
  build-time edit-event side effect; `PNV-00004` was not posted; production
  posted neither. A one-line policy decision (post both / remove r15) is required.

## Evidence index

- `cdj_deletion_ledger.csv` — this directory; 159 signature groups, 171 removed-class rows.
- Independent scripts (temp dir, read-only): `recon_gl.py`,
  `recon_fix2.py` (alias-aware), `crosscheck.py`, `sundry_map.py`,
  `legs.py`, `gate2_lineage.py`, `gate2_evidence.py`, `gate2_resid.py`,
  `verify_fix.py`; outputs `recon_fix.txt`, `recon_fix2.txt`,
  `crosscheck_fix.txt`, `gate2_evidence_summary.txt`, `gate2_resid_out.txt`.
- Workbook-internal evidence: `GL_AUDIT` rows 3–9 (incl. the candidate's
  2026-10-14 HOLD row with the full 150-line unmapped detail).
- VBA exports: `gate2_work/vba_baseline/` (shipped), `gate2_work/vba_patched/`
  (intended patch), `vba_candidate/` in temp (extracted from the candidate).
