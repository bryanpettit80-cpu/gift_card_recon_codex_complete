# Codex Prompts

## Build or repair the repo

```text
You are working in a Python repo that automates restaurant gift card reconciliation.

Business objective:
Reconcile one monthly gift card summary, multiple weekly gift card activity reports, and POS control totals. The output must be an Excel workbook suitable for accounting review.

Guardrails:
- Deterministic Python math only.
- Do not use AI judgment for reconciliation values.
- Keep POS controls on the Reconciliation tab.
- Do not add extra-week, calendar cutoff, or period-boundary commentary unless explicitly requested.
- Preserve the tabs: Reconciliation, Weekly Activity Detail, Daily Activity Detail, Raw Detail, Source Files, Exception Log.

Known May 2026 store 9354 controls:
- Weekly activity activations total: 11507.00
- Weekly activity redemptions total: 49867.48
- POS gift card issue: 11642.00
- POS gift card payment: 49869.75
- POS issue variance: 135.00
- POS payment variance: 2.27
- POS net variance: 132.73

Run pytest. Fix failures. Explain only material changes.
```

## Add a new feature

```text
Add the requested feature to the gift card reconciliation repo.
Before editing, identify the parser, reconciliation, writer, or PowerShell layer affected.
After editing, run pytest and report the impact on the May 2026 store 9354 expected values.
Do not change the existing reconciliation math unless a failing test proves it is wrong.
```
