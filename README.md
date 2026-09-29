# globalsmile

A local Git repository initialized for the globalsmile project.

## Getting started

```bash
git status
```

# v8.2 GL Automation Final Review

**Static precheck: PASS**

- 46 GL accounts × 12 months = 552 rows
- CDJ + CRJ + GJ posting sources
- PJ / PJ Non-Vat / SJ excluded as feeder journals 

## Code fixes

1. modGLAggregation.bas no longer exposes duplicate RefreshGL / RefreshAllGL.
2. modGLRefresh.bas is the canonical refresh entrypoint.
3. The posting matrix includes CDJ + CRJ + GJ.
4. The five documented v8.1 GJ mappings are executable at runtime.
5. CDJ sundry residual logic no longer double-counts the cash-out amount.
6. GL_V8_CALC writes cumulative ending balances.

## Runtime gate

The uploaded v8.1-rebuild workbook preserves its existing vbaProject.bin, so the generated review workbook is **not presented as runtime-certified**.

Desktop Excel gate: import the four v8.2 BAS modules, compile, run RefreshAllGL(2026), and verify GL_AUDIT plus the 552-row GL_V8_CALC.
