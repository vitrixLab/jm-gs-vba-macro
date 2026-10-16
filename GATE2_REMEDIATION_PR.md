# Gate 2 remediation PR — CRJ month partition fix

## Change made in this PR

modGLScanEngine.bas now attributes CRJ invoice rows to the month identified by the enclosing monthly TOTAL row (month label in column D; TOTAL marker in column F), instead of using column O.

The Gate 2 investigation found that column O contains SJ import/log timestamps, not reliable transaction dates. That caused 559 invoices to be counted as invalid and three January-partition invoices to be attributed to August. The new source routine:

- scans CRJ monthly TOTAL boundaries;
- validates month labels are valid and strictly increasing;
- aggregates invoice rows into the partition month;
- fails closed on invalid/out-of-order partitions or invoice rows outside a monthly partition;
- retains the existing CRJ amount columns and GL account mappings.

## Still blocking workbook promotion

This is a source-level remediation, not a completed workbook repair. The repository connector can update text sources, but cannot upload or rewrite the XLSM binary through its contents API. The source change must be integrated into a candidate workbook and tested in desktop Excel before the workbook can be promoted.

The following blockers remain and are deliberately not guessed:

1. Seven accounting classifications still need the accounting owner to choose the correct existing COA account or authorize a COA change: Representation; Partners Clinicians Fee; SSS,HDMF AND PHIC ER SHARE; Partners Clinicians Fee-Corporators; Incentives; Outside Services; Training and Seminar.
2. Candidate data integrity: restore the 132 PJ column-C day cells from the unchanged production baseline; independently verify the CDJ deletion ledger, including the four blank-particulars rows marked NO_TWIN_IN_CANDIDATE; resolve the accidental Non-Vat row PNV-00002 against its sibling PNV-00004.
3. No sheet redesign: preserve the existing 17 sheet names and the existing GL 46-account × 12-month layout. Do not add or remove sheets as part of this fix.
4. Excel runtime acceptance: compile VBA; run the gate; confirm Debit = Credit within ₱0.01, unmapped = ₱0, invalid = 0; compare all 46 accounts and 552 month cells to independent journal aggregation (0 differences); run a second refresh and prove deterministic output; verify the production workbook SHA remains unchanged.

## Evidence and status

- Baseline reviewed: c302898f8a0b847a487447414b93129e39e2f898.
- Existing Gate 2 result remains HOLD: Debit ₱4,431,497.52; Credit ₱4,972,059.43; difference −₱540,561.91; unmapped ₱698,740.35; invalid 559.
- The numbers above are the pre-patch evidence report. They are not claimed as post-patch test results.
- Desktop Excel compile/runtime, live GL write-back, independent 46×12 reconciliation, and second-refresh determinism are NOT RUN in this PR.

## Acceptance rule

Do not merge or promote this as a certified workbook fix until the owner classifications are documented, source data integrity is restored and verified, the ledger classifications reconcile, and all runtime acceptance checks pass.
