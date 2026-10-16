# Gate 2 remediation PR — CRJ month partition fix

## Change made in this PR

modGLScanEngine.bas now attributes CRJ invoice rows to the month identified by the enclosing live monthly boundary row (month label in column D; columns F and G blank), instead of using column O.

> Correction note: the original merged version of this change detected boundaries by `F = "TOTAL"`. Independent review of merge `19e6d50` showed the live CRJ monthly boundary rows (r115, r197, r317, r390, r487, r582) have a blank column F; the only `F = "TOTAL"` cells are a legacy v8.3.3 summary block at rows 2690-2783. Against the real sheet the old detector therefore posted all 562 live invoices into January via that stale block with inv=0 (gate-invisible wrong-month result). This branch replaces the marker test with the live-boundary predicate (month in D + F blank + G blank), validated across all 3,384 CRJ rows: exactly 6 detections, zero false positives. Dead code left by the earlier rewrite (unused `n`, unused `EntryLogDate`) is removed.

The Gate 2 investigation found that column O contains SJ import/log timestamps, not reliable transaction dates. That caused 559 invoices to be counted as invalid and three January-partition invoices to be attributed to August. The new source routine:

- scans CRJ monthly boundary rows (month label in column D, columns F and G blank);
- scans through the last month label in column D so a trailing boundary row (blank in F/G) is never cut out of the scan range;
- validates the first boundary is January and each subsequent boundary is exactly the next month (strictly contiguous, no gaps, duplicates, or reordering);
- aggregates invoice rows into the enclosing boundary's month;
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

- Baseline reviewed: c302898f8a0b847a487447414b93129e39e2f898; merged PR #6 reviewed at 19e6d50223bd5816bad221e5003bdb7ffc801eaf.
- This branch re-simulates the corrected detector against Global-Smile_2026-v8.4-fix.xlsm (read-only): 6 boundaries detected at r115/197/317/390/487/582; 562 invoices posted Jan-Jun as 593,273.12 / 450,514.28 / 837,607.99 / 421,057.12 / 548,191.71 / 606,743.55 (sum 3,457,387.77 = boundary row H totals, exact to the cent); legacy block rows 2690-2783 excluded; inv=0 only because no invoice lies outside a partition.
- Existing Gate 2 result remains HOLD: Debit ₱4,431,497.52; Credit ₱4,972,059.43; difference −₱540,561.91; unmapped ₱698,740.35; invalid 559.
- The numbers above are the pre-patch evidence report. They are not claimed as post-patch test results.
- Desktop Excel compile/runtime, live GL write-back, independent 46×12 reconciliation, and second-refresh determinism are NOT RUN in this PR.

## Acceptance rule

Do not merge or promote this as a certified workbook fix until the owner classifications are documented, source data integrity is restored and verified, the ledger classifications reconcile, and all runtime acceptance checks pass.
