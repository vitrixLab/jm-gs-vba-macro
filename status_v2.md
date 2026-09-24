# Project Status — Global-Smile v8.0 (Evidence-Reconciled POC)

## Overall Status: **POC — Reconciliation PASS / COA Mapping HOLD**

This status replaces the earlier `status.md` certification wording.

The v8.0 workbook has a verified balanced posting population and a 46-account × 12-month GL structure. However, the available workbook evidence does **not** support declaring full COA coverage or full seven-gate certification. COA mapping remains on HOLD until the outstanding source labels are explicitly classified into the 46-account COA.

---

## Evidence Basis

**Workbook tested:** `Global-Smile_2026-v8.0-FINAL.xlsm`

**Local SHA-256:**

`94de52d34a900996c4283b9261bb866be86e8c3e30331964f082cf641060b913`

**VBA project SHA-256:**

`85561d8eac34e24f43266cf7e521a997cc22f5b7691683571da2e389615b6639`

The workbook opens as an XLSM package and contains the VBA project. Workbook structure was inspected directly rather than inferred from the status document.

---

## Verified Source Reconciliation

| Source | Debit | Credit | Difference | Result |
|---|---:|---:|---:|---|
| CDJ | 4,500.00 | 4,500.00 | 0.00 | PASS |
| CRJ | 13,240.00 | 13,240.00 | 0.00 | PASS |
| GJ | 238,405.75 | 238,405.75 | 0.00 | PASS |
| **Combined** | **256,145.75** | **256,145.75** | **0.00** | **PASS** |

These figures independently reconcile to zero difference.

### Interpretation

The source journals are balanced.

This does **not**, by itself, prove that every posting label has been successfully mapped to one of the 46 GL accounts.

---

## GL Matrix

The workbook contains **46 GL account titles** with a 12-month structure:

**46 × 12 = 552 account-month records**

This structural claim is verified.

---

## COA Coverage Status

### Result: **HOLD — not fully resolved**

The earlier v8.0 status claimed that all previous unmapped activity totaling **6,154.00** had been resolved through GJ aliases.

The workbook evidence does not support that conclusion.

The following source labels remain outside the verified 46-account GL title set:

| Source label | Amount | Verified status |
|---|---:|---|
| Medical Equipment | 2,992.50 credit | NOT VERIFIED AS A 46-COA ACCOUNT |
| Cost of Revenue | 1,139.00 debit | **UNMAPPED IN GL_AUDIT** |
| Supplies | 2,000.00 debit | NOT VERIFIED AS A 46-COA ACCOUNT |
| Bank Charge | 15.00 debit | NOT VERIFIED AS A 46-COA ACCOUNT |
| Charges | 7.50 debit | NOT VERIFIED AS A 46-COA ACCOUNT |
| **Total** | **6,154.00** | **COA MAPPING HOLD** |

In particular, the workbook's `GL_AUDIT` contains an entry identifying:

`GJ | Cost of Revenue | 1139 | UNMAPPED - ACCOUNT LIST REQUIRES REVIEW`

Therefore the statement that all five labels are already resolved must not be treated as certified evidence.

---

## Workbook Structure Check

The workbook contains these sheets:

- PJ
- PJ Non-Vat
- CDJ
- SJ
- CRJ
- GJ
- ZZZ_CompileProbe
- GL
- SUPPLIERS DATA
- CRJ2
- PJ2
- PJ Non-Vat2
- CDJ old
- SJ2
- GL_AUDIT

### Important correction

There is **no `GL_V8_CALC` worksheet** in the tested workbook.

Therefore the previous statement:

> GL_V8_CALC sheet built with 552 rows

is **not verified and is removed from the completion claims**.

---

## Gate Evidence Matrix

| Gate | Claim | Evidence status |
|---|---|---|
| G1 | CDJ/CRJ/GJ mappings | **UNVERIFIED** as a complete engine-level claim |
| G1 | Feeder journals excluded from posting population | **UNVERIFIED** by execution test |
| G2 | Source-specific month/date extraction | **UNVERIFIED** by execution test |
| G2 | Invalid numeric/date count = 0 | **UNVERIFIED** by execution test |
| G3 | Combined debit = credit | **PASS — 256,145.75 = 256,145.75** |
| G4 | 46-account COA coverage | **HOLD — not fully resolved** |
| G5 | 46 × 12 matrix | **PASS — 552 records structurally verified** |
| G6 | Existing GL protected while HOLD | **UNVERIFIED by runtime execution** |
| G7 | Graphify separated from accounting | **UNVERIFIED by complete source/runtime test** |

### Gate conclusion

The evidence supports **G3 and G5**.

The evidence does **not** support declaring all seven gates PASS.

---

## GL Protection

The intended design is:

- Existing GL remains protected while validation is HOLD.
- GL writes should remain blocked until validation returns PASS.

These are **design claims requiring runtime VBA execution testing**. They are not marked PASS solely from workbook inspection.

`GL_AUDIT` is present in the workbook and contains audit information.

---

## Implementation Claims

The earlier status listed:

- `modGLGate.bas`
- `modGLAggregation.bas`
- `modGLWorkbookMap.bas`
- `modEngine.bas`
- `modPJAutomation.bas`
- Graphify modules

The presence and exact line counts of these VBA modules are **not treated as independently certified by this status document** unless verified from the VBA project/source extraction.

The VBA project is present in the XLSM package, but package presence alone is not equivalent to runtime certification.

---

## What Is Certified by This Status

### Certified from workbook evidence

- CDJ gross debit/credit reconciliation
- CRJ gross debit/credit reconciliation
- GJ gross debit/credit reconciliation
- Combined debit/credit reconciliation
- Zero combined difference
- 46-account GL structure
- 12-month GL structure
- 552 account-month records
- Presence of `GL_AUDIT`

### Not certified

- Complete 46-account COA coverage
- Resolution of all 6,154.00 previously unmapped labels
- Existence of `GL_V8_CALC`
- All seven gates PASS
- Runtime protection behavior
- Runtime GL write authorization behavior
- Complete VBA module-level runtime behavior
- Production readiness

---

## Required Next Gate

### Accounting / COA Resolution

The five source labels totaling **6,154.00** must be explicitly mapped to the existing 46-account COA or the 46-account COA must be formally revised.

No automatic alias should be treated as an accounting authorization unless the resulting GL account classification is explicitly confirmed.

### Runtime Validation

After COA resolution:

1. Execute the v8.0 VBA engine.
2. Verify source-specific extraction.
3. Verify unmapped-account count reaches zero.
4. Verify `ValidateGLConsistency` behavior.
5. Verify GL write protection during HOLD.
6. Verify GL output against independently calculated source totals.
7. Re-run the gate matrix and record actual PASS evidence.

---

## Status Decision

**Current state:**

`POC — Reconciliation PASS / COA Mapping HOLD`

**Not yet:**

`COMPLETE — POC Certified`

The distinction is intentional: balanced source journals and a correctly sized GL matrix are proven, while complete COA mapping and runtime certification remain to be demonstrated.

---

**Project Lead:** vitrixLab  
**Version:** v8.0  
**Mode:** POC — Evidence-Reconciled  
**Date:** 2026-09-25
