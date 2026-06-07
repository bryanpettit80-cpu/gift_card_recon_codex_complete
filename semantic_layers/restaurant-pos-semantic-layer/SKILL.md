---
name: restaurant-pos-semantic-layer
description: Use when analyzing Ruth's Chris restaurant POS exports, Micros 3700 text exports, gift card reconciliation inputs, tender detail, daily POS controls, discounts, menu mix, revenue centers, or labor/timecard data from this workspace.
---

# Restaurant POS Semantic Layer

Use this context layer before answering questions about Micros 3700 exports, restaurant POS controls, tender activity, gift card payments/issues, menu mix, discounts, or timecard/labor exports in this workspace.

## Source Order

1. Prefer current user-provided files and exports for the period being analyzed.
2. Use `references/source-inventory.md` to identify available local sources and known gaps.
3. Use `references/micros3700-export-map.md` to interpret the known `Micros3700.7z` export shape.
4. Use the existing `gift-card-reconciliation` skill for workbook reconciliation math and output rules.

## Core Rules

- Treat Micros files as CSV-style text rows with single-quoted text values and no header row.
- Do not infer official Micros column names for wide control tables unless a schema/export definition is provided.
- For gift card reconciliation, start with `TENDER_DETAIL.TXT` and tender name `G C Payment` for redemptions/payments.
- Treat employee files (`EMPDEF.TXT`, `EMPTL.TXT`, `TIMECARD.TXT`, `JOBRTDEF.TXT`) as sensitive personnel/labor data. Summarize only what is needed.
- Keep POS controls on the Reconciliation tab when generating reconciliation workbooks. Do not create extra POS tie-out tabs unless explicitly requested.
- When sources disagree, report the controlling source, the variance, and the exact file/date/tender/category involved.

## Common Outputs

- Gift card payment totals by business date
- Tender summary and variance checks
- Discount totals by reason
- Menu sales by item, family group, or major group
- Daily POS control tie-outs
- Labor/timecard summaries when requested

