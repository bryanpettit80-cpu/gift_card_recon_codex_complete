---
name: gift-card-reconciliation
description: Use when reconciling restaurant gift card summary, weekly activity, and POS control files into a formatted Excel workbook.
---

When triggered:

1. Confirm the expected input folder exists.
2. Confirm exactly one gift card summary file exists.
3. Confirm one or more gift card activity files exist.
4. Confirm POS controls exist either as `pos_controls.csv` or command-line values.
5. Run the reconciliation script.
6. Run pytest.
7. Review the output workbook tabs.
8. Report only the primary variances:
   - Summary vs activity activations
   - Summary vs activity redemptions
   - POS Gift Card Issue variance
   - POS Gift Card Payment variance
   - Net Gift Card Impact variance

Business-rule guardrails:

- Keep POS controls on the Reconciliation tab.
- Do not create a separate POS tie-out tab unless explicitly requested.
- Do not mention extra weeks, calendar cutoffs, or period-boundary theories unless explicitly requested.
- Do not change reconciliation math without updating tests.
