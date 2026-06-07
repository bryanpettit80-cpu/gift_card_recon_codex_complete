# Gift Card And POS Reconciliation Rules

These rules complement the workspace `gift-card-reconciliation` skill.

## Primary Tie-Outs

Use these checks when the user asks for gift card reconciliation:

1. Summary vs activity activations
2. Summary vs activity redemptions
3. POS Gift Card Issue variance
4. POS Gift Card Payment variance
5. Net Gift Card Impact variance

## Micros POS Inputs

- Gift card payments/redemptions: start with `TENDER_DETAIL.TXT` rows where tender name is `G C Payment`.
- Tender totals: aggregate `TENDER_DETAIL.TXT` by date, tender number, tender name, and status.
- Discounts: aggregate `DISCOUNT_DETAIL.TXT` by date and reason.
- Control totals: use `DLYSYSTT.TXT`, `DLYSYSTL.TXT`, `DLYRVCTT.TXT`, and `DLYRVCTL.TXT` only after column mapping or total matching is clear.

## Workbook Guardrails

- Keep POS controls on the Reconciliation tab.
- Do not create a separate POS tie-out tab unless explicitly requested.
- Do not change reconciliation math without updating tests.
- Do not mention extra weeks, calendar cutoffs, or period-boundary theories unless explicitly requested.

## Variance Reporting

When reporting variances, include:

- Period or business date
- Source file or workbook tab
- Metric name
- Source value
- Compared value
- Variance amount
- Whether the result is exact, inferred, or blocked by missing schema/source data

