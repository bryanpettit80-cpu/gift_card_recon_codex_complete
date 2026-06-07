# Micros 3700 Export Map

This map reflects the observed `Micros3700.7z` export inspected on 2026-06-05. The files have no header rows, so field meanings are inferred from filenames, row shapes, values, and totals.

## Summary

| File | Rows | Columns | Observed Date Range | Use |
|---|---:|---:|---|---|
| `DISCOUNT_DETAIL.TXT` | 759 | 3 | 2025-05-01 to 2026-06-04 | Discount totals by date/reason |
| `DLYRVCTL.TXT` | 999 | 59 | 2025-05-01 to 2026-06-04 | Revenue-center daily control totals |
| `DLYRVCTT.TXT` | 999 | 133 | 2025-05-01 to 2026-06-04 | Revenue-center tracking/tender totals |
| `DLYSYSTL.TXT` | 401 | 74 | 2025-05-01 to 2026-06-05 | System daily control totals |
| `DLYSYSTT.TXT` | 351 | 132 | 2025-05-01 to 2026-06-04 | System tracking/tender totals |
| `EMPDEF.TXT` | 364 | 180 | 2001-01-01 to 2026-06-02 | Employee definitions; sensitive |
| `EMPTL.TXT` | 7,175 | 51 | 2025-05-01 to 2026-06-05 | Employee/labor totals; sensitive |
| `FAMILY_GROUP_DETAIL.TXT` | 4,965 | 17 | 1997-06-19 to 2026-06-04 | Sales by family group |
| `JOBDEF.TXT` | 14 | 44 | 1997-07-06 to 2026-01-03 | Job definitions |
| `JOBRTDEF.TXT` | 888 | 12 | 2007-03-19 to 2026-06-02 | Job rate definitions; sensitive |
| `MAJOR_GROUP_DETAIL.TXT` | 1,666 | 19 | 1997-06-18 to 2026-06-04 | Sales by major group |
| `MENU_ITEM_DETAIL.TXT` | 57,799 | 6 | 2025-05-01 to 2026-06-04 | Menu item quantity/sales |
| `RESTDEF.TXT` | 1 | 167 | 1997-12-09 to 2005-01-01 | Restaurant definition |
| `RESTSTATUS.TXT` | 1 | 14 | 1997-12-12 to 2026-06-05 | Restaurant/Micros status |
| `RVCDEF.TXT` | 4 | 224 | 1997-12-09 | Revenue center definitions |
| `TENDER_DETAIL.TXT` | 2,644 | 5 | 2025-05-01 to 2026-06-04 | Tender totals by date/name |
| `TIMECARD.TXT` | 940 | 48 | 2026-05-05 to 2026-06-05 | Timecard/labor detail; sensitive |
| `TRKGRDEF.TXT` | 20 | 6 | none observed | Tracking group definitions |
| `TRKTTDEF.TXT` | 1,257 | 6 | none observed | Tracking total definitions |

## Key Inferred Tables

### `TENDER_DETAIL.TXT`

Observed shape:

```text
date, amount, tender_number, tender_name, status
```

Important observed totals:

- Total tender amount: `$9,665,888.34`
- Visa: `$4,496,789.04`
- Amex: `$1,664,484.18`
- Master Card: `$1,280,260.01`
- Tips Paid: `$1,137,930.56`
- Cash: `$538,979.98`
- G C Payment: `$397,264.54`
- Discover: `$108,983.89`

Status values observed:

- `T`: `$8,527,957.78`
- `P`: `$1,137,930.56`
- `S`: `$0.00`

Gift card reconciliation starting point:

- Use tender name `G C Payment` as POS gift card payment/redemption activity unless another gift-card tender is explicitly provided.

### `DISCOUNT_DETAIL.TXT`

Observed shape:

```text
date, amount, discount_reason
```

Observed total discount amount: `-$95,250.03`

Observed reason totals:

- Promo Disc: `-$65,983.85`
- 15% Discount: `-$12,924.35`
- ADV Comp: `-$7,995.25`
- Emp Meal: `-$5,100.14`
- Donation: `-$3,246.44`

### `MENU_ITEM_DETAIL.TXT`

Observed shape:

```text
date, quantity, sales_amount, menu_item_name, major_group, family_group
```

Observed total sales: `$6,574,949.80`
Observed total quantity: `290,052`

Top observed items by sales:

- RIBEYE 16oz: `$374,904.00`
- PETITE FILET: `$301,792.00`
- FILET: `$276,396.00`
- STUFF CHICKEN: `$227,332.00`
- RC PF/SHR: `$183,568.00`
- NY STRIP: `$173,466.00`
- LAMB CHOPS: `$157,248.00`
- COWBOY: `$150,634.00`

### `MAJOR_GROUP_DETAIL.TXT`

Useful inferred fields include date, major group identifiers, quantity, amount, and major group name.

Observed total sales: `$6,574,949.80`
Observed major group sales:

- Food: `$5,332,206.04`
- Liquor: `$636,677.81`
- Wine: `$468,574.21`
- Card Items: `$88,310.92`
- Beer: `$31,923.50`
- Banquet AV/RC: `$17,150.00`
- Retail: `$107.32`

### `FAMILY_GROUP_DETAIL.TXT`

Useful inferred fields include date, family group identifiers, quantity, amount, and family group name.

Observed total sales: `$6,574,949.80`
Top observed family group sales:

- Meat: `$2,828,572.50`
- Other Food: `$1,325,038.78`
- Liquor: `$618,788.31`
- Seafood Apps: `$584,948.75`
- Seafood: `$252,950.00`
- Glass Wine: `$250,091.00`
- Bottled Wine: `$218,483.21`
- Lobster Apps: `$158,099.00`

### Control Tables

`DLYRVCTL.TXT`, `DLYRVCTT.TXT`, `DLYSYSTL.TXT`, and `DLYSYSTT.TXT` are wide daily control/tracking tables. Use them for POS control tie-outs after mapping columns from a Micros schema, report definition, or matching totals to `TRKTTDEF.TXT`.

Do not label wide-table columns as official business metrics without a schema.

### Definition Tables

- `RESTDEF.TXT`: restaurant definition, including Ruth's Chris and Virginia Beach location fields.
- `RESTSTATUS.TXT`: restaurant status/Micros version. Version observed: `V5.5`.
- `RVCDEF.TXT`: revenue centers observed as Dining Room, Bar, Takeout, Banquets.
- `JOBDEF.TXT`: job names observed include Training, Host, Runner, Bartender, Server, Kitchen, SA, Supervisor, Server Coach, Runner Coach, SA Coach, Patio Server, Bar SA, Lead Host.
- `TRKGRDEF.TXT` and `TRKTTDEF.TXT`: tracking group/total definitions, useful for interpreting control table positions.

### Sensitive Labor Tables

- `EMPDEF.TXT`: employee master/definition data.
- `EMPTL.TXT`: employee/labor totals.
- `JOBRTDEF.TXT`: job rate assignments.
- `TIMECARD.TXT`: time punches and pay/hour-related fields.

Observed `TIMECARD.TXT` summary:

- Rows: `940`
- Unique employees: `57`
- Total hours column observed: `5,498.64`
- Total pay-like column observed: `$51,974.46`
- Date range: 2026-05-05 to 2026-06-05

Use aggregated labor summaries by default.

