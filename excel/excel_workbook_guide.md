# Excel analytical workbook — build guide

How to build the standalone Excel deliverable from
`data/processed/luxury_retail_clienteling_analytics.xlsx`.

Every formula below references the real column letters and row counts produced
by `python -m src.run_pipeline` at seed 42. If you regenerate with different
parameters, the row counts move but the column letters do not.

## Sheet map

| Sheet | Rows (incl. header) | Grain |
|---|---|---|
| `Transactions` | 15,018 | One order line |
| `Customer Master` | 3,401 | One client |
| `Product Master` | 241 | One SKU |
| `Boutique Master` | 11 | One boutique |
| `Associates` | 57 | One associate |
| `Clienteling Activities` | 3,973 | One interaction |
| `Marketing Campaigns` | 27 | One campaign |

### Column letters

**`Transactions`**

| Col | Field | Col | Field |
|---|---|---|---|
| A | transaction_id | K | discount_pct |
| B | order_id | L | unit_net_price |
| C | transaction_date | M | gross_amount |
| D | customer_id | N | net_amount |
| E | sku | O | unit_cost |
| F | boutique_id | P | is_return |
| G | associate_id | Q | is_full_price |
| H | channel | R | interaction_id |
| I | quantity | S | campaign_id |
| J | unit_list_price | | |

**`Customer Master`**

| Col | Field | Col | Field |
|---|---|---|---|
| A | customer_id | G | client_tier |
| B | client_name | H | country |
| C | join_date | I | region |
| D | acquisition_source | J | email_opt_in |
| E | preferred_boutique_id | K | churn_date |
| F | consent_status | L | primary_category_affinity |

**`Clienteling Activities`**

| Col | Field |
|---|---|
| A | interaction_id |
| B | customer_id |
| C | associate_id |
| D | boutique_id |
| E | interaction_date |
| F | interaction_type |
| G | converted |
| H | outcome |

## Step 0 — convert each sheet to a Table

Select any cell in the data, press `Ctrl+T`, tick "My table has headers", then
rename on the Table Design tab:

| Sheet | Table name |
|---|---|
| Transactions | `tblTransactions` |
| Customer Master | `tblCustomers` |
| Product Master | `tblProducts` |
| Boutique Master | `tblBoutiques` |
| Associates | `tblAssociates` |
| Clienteling Activities | `tblClienteling` |
| Marketing Campaigns | `tblCampaigns` |

Do this before writing any formula. Structured references such as
`tblTransactions[net_amount]` survive new rows; `$N$2:$N$15018` does not, and
every figure in the workbook quietly goes stale the next time the data is
refreshed.

Both forms are given below: structured references for the real build, A1
references so you can verify a formula against a specific cell.

Add one helper column to `tblTransactions` — it makes every downstream
formula shorter and is the single rule the whole workbook depends on:

```excel
// Column T, header: net_sales_value
=IF([@is_return]=TRUE, 0, [@net_amount])

// A1 equivalent, T2 filled down to T15018
=IF(P2=TRUE,0,N2)
```

**Everything monetary in this workbook sums column T, never column N.** That
is what "excluding returns" means in practice.

Also add an order-month column for the trend charts:

```excel
// Column U, header: order_month
=EOMONTH([@transaction_date],0)

// A1 equivalent
=EOMONTH(C2,0)
```

Format column U as `mmm-yy`.

---

## Tab 1 — Executive Dashboard

### PivotTable 1: monthly revenue trend

1. Select `tblTransactions` → Insert → PivotTable → New worksheet, name it
   `pvt_Revenue`.
2. Drag `order_month` to **Rows**. Right-click → Group → ungroup if Excel
   auto-groups into Years/Quarters; the column is already month-ended.
3. Drag `net_sales_value` to **Values**, set to Sum, rename to `Net Sales`.
4. Drag `order_id` to **Values**, set to **Distinct Count** (requires
   "Add this data to the Data Model" when the PivotTable is created), rename to
   `Transactions`.
5. Add a calculated field for ATV. Note that Distinct Count is unavailable in
   classic calculated fields, so compute ATV beside the pivot instead:

```excel
// In the cell to the right of the first data row
=GETPIVOTDATA("Net Sales",$A$3,"order_month",A4)
 / GETPIVOTDATA("Transactions",$A$3,"order_month",A4)
```

6. Insert → PivotChart → Combo. Net Sales as clustered column on the primary
   axis, ATV as a line on the secondary axis. Format the secondary axis to
   currency with no decimals.

### PivotTable 2: repeat purchase rate

Distinct-count-of-clients-with-2+-orders cannot be done in a single pivot.
Build it on the RFM tab (below) and reference the result:

```excel
// Repeat Purchase Rate
=COUNTIFS(RFM!D:D,">=2") / COUNTA(RFM!A:A)-1
```

Better, avoiding the header-offset trap entirely:

```excel
=COUNTIF(tblRFM[frequency],">=2") / ROWS(tblRFM[frequency])
```

### KPI block

Place these above the charts. All reference the full dataset; wire them to
slicers afterwards.

```excel
// Net Sales
=SUM(tblTransactions[net_sales_value])
// A1: =SUMIF(Transactions!$P$2:$P$15018,FALSE,Transactions!$N$2:$N$15018)

// Transaction count (distinct orders, excluding fully returned orders)
=SUMPRODUCT((tblTransactions[net_sales_value]>0)
           /COUNTIFS(tblTransactions[order_id],tblTransactions[order_id]&"",
                     tblTransactions[net_sales_value],">0"))

// Average Transaction Value
=SUM(tblTransactions[net_sales_value]) / [the cell above]

// Units sold
=SUMIFS(tblTransactions[quantity],tblTransactions[is_return],FALSE)

// Average Unit Retail
=SUM(tblTransactions[net_sales_value])
 / SUMIFS(tblTransactions[quantity],tblTransactions[is_return],FALSE)

// Full-Price Sales Share
=SUMIFS(tblTransactions[net_amount],
        tblTransactions[is_return],FALSE,
        tblTransactions[is_full_price],TRUE)
 / SUM(tblTransactions[net_sales_value])

// Return rate by value
=SUMIFS(tblTransactions[net_amount],tblTransactions[is_return],TRUE)
 / SUM(tblTransactions[net_amount])
```

The distinct-order `SUMPRODUCT` above is the classic Excel workaround and it is
slow on 15,000 rows. If the workbook will be used regularly, load the table
into the Data Model (Power Pivot) and use `DISTINCTCOUNT` instead — the DAX in
`bi/measures.dax` works unchanged in Power Pivot.

### Slicers

Insert → Slicer on `tblTransactions`: `boutique_id`, `channel`, and
`order_month` as a Timeline. Then Report Connections → tick every PivotTable so
one slicer drives the whole dashboard.

---

## Tab 2 — RFM & CLV

Create a sheet named `RFM`. Build the client list first, then layer the metrics.

### Column A — unique client list

```excel
// A1 header: customer_id, A2:
=SORT(UNIQUE(FILTER(tblTransactions[customer_id],
                    tblTransactions[is_return]=FALSE)))
```

This spills down automatically. Only clients with a non-returned purchase
appear, which is correct — a client whose single order was returned has no RFM
profile.

### Column B — client name

```excel
=XLOOKUP([@customer_id],tblCustomers[customer_id],tblCustomers[client_name],"Unknown")

// A1 equivalent, B2 filled down
=XLOOKUP(A2,'Customer Master'!$A$2:$A$3401,'Customer Master'!$B$2:$B$3401,"Unknown")
```

### Column C — Recency (days since last purchase)

Set the snapshot date in a named cell first. Put `2026-09-30` in `Config!$B$1`
and name it `SnapshotDate`. Hard-coding `TODAY()` here would make every number
in the workbook change overnight and stop matching the Python output.

```excel
=SnapshotDate - MAXIFS(tblTransactions[transaction_date],
                       tblTransactions[customer_id],[@customer_id],
                       tblTransactions[is_return],FALSE)

// A1 equivalent, C2 filled down
=$B$1 - MAXIFS(Transactions!$C$2:$C$15018,
               Transactions!$D$2:$D$15018,A2,
               Transactions!$P$2:$P$15018,FALSE)
```

Format as Number, 0 decimals — Excel will otherwise inherit the date format and
show `1900-03-12` instead of `71`.

### Column D — Frequency (count of distinct orders)

```excel
=SUMPRODUCT((tblTransactions[customer_id]=[@customer_id])
           *(tblTransactions[is_return]=FALSE)
           /COUNTIFS(tblTransactions[order_id],tblTransactions[order_id]))

// A1 equivalent, D2 filled down
=SUMPRODUCT((Transactions!$D$2:$D$15018=A2)
           *(Transactions!$P$2:$P$15018=FALSE)
           /COUNTIFS(Transactions!$B$2:$B$15018,Transactions!$B$2:$B$15018))
```

Why the division: `COUNTIFS` returns how many lines share each order_id, so
dividing by it makes a 3-line order contribute `1/3 + 1/3 + 1/3 = 1`. Counting
lines instead of orders is the single most common error in an RFM sheet, and it
inflates Frequency for exactly the clients who buy multi-piece baskets — the
ones you least want to misread.

A faster alternative on large files, if you accept a helper column: add
`is_first_line_of_order` to `tblTransactions` —

```excel
=IF(COUNTIFS($B$2:B2,[@order_id])=1,1,0)
```

— then Frequency is a plain `SUMIFS` over it.

### Column E — Monetary (sum of net spend)

```excel
=SUMIFS(tblTransactions[net_amount],
        tblTransactions[customer_id],[@customer_id],
        tblTransactions[is_return],FALSE)

// A1 equivalent, E2 filled down
=SUMIFS(Transactions!$N$2:$N$15018,
        Transactions!$D$2:$D$15018,A2,
        Transactions!$P$2:$P$15018,FALSE)
```

### Columns F, G, H — RFM scores (1–5)

Quintiles via `PERCENTRANK`. Recency is inverted, because fewer days is better:

```excel
// F — R score
=6 - MIN(5, INT(PERCENTRANK.INC($C$2:$C$2600,[@Recency])*5)+1)

// G — F score
=MIN(5, INT(PERCENTRANK.INC($D$2:$D$2600,[@Frequency])*5)+1)

// H — M score
=MIN(5, INT(PERCENTRANK.INC($E$2:$E$2600,[@Monetary])*5)+1)
```

Adjust `2600` to the last row the spill in column A actually reaches (2,470 at
seed 42 — leave headroom). The `MIN(5, ...)` caps the top value: the highest
observation has a percent rank of exactly 1.0, which would otherwise score 6.

### Column I — Segment

Precedence matters. A high-value client who has gone quiet must be flagged as
lapsing before any other rule can claim them:

```excel
=IFS(
    AND([@[M score]]>=4, [@Recency]>180),                      "Lapsing High-Value",
    AND([@[M score]]>=4, [@[F score]]>=4, [@Recency]<=180),    "High-Value Loyal",
    AND([@Tenure]<=365, [@Recency]<=180),                      "Promising New",
    TRUE,                                                       "Infrequent"
)
```

Tenure (column J) is days since the client's first purchase:

```excel
=SnapshotDate - MINIFS(tblTransactions[transaction_date],
                       tblTransactions[customer_id],[@customer_id],
                       tblTransactions[is_return],FALSE)
```

### Columns K–N — 12-month CLV

```excel
// K — Average Order Value
=[@Monetary]/[@Frequency]

// L — Expected orders in the next 12 months
=MIN(12, [@Frequency] / MAX(0.5, [@Tenure]/365.25))

// M — Retention rate (base rate x recency factor)
=MEDIAN(0.05, BaseRetention * XLOOKUP([@[R score]],{1;2;3;4;5},{0.35;0.6;0.85;1.05;1.2}), 0.95)

// N — CLV (12-month gross margin, discounted)
=[@AOV] * [@[Expected Orders]] * [@Retention] * GrossMargin / (1+DiscountRate)^0.5
```

Three named cells on a `Config` sheet drive this:

| Name | Value | Source |
|---|---|---|
| `BaseRetention` | 0.6005 | 12-month repeat rate from the cohort model |
| `GrossMargin` | 0.6937 | realised blended margin |
| `DiscountRate` | 0.10 | assumed cost of capital |

`MEDIAN(low, value, high)` is the compact way to clamp a number between two
bounds — shorter and less error-prone than nested `MIN`/`MAX`.

### Segment summary PivotTable

Insert a PivotTable on `tblRFM`: `Segment` to Rows, `customer_id` to Values as
Count, `Monetary` to Values as Sum, `CLV` to Values as Sum. Add a
`% of Parent Total` version of Monetary — that one column is what turns the tab
from a list into an argument.

---

## Tab 3 — Clienteling ROI

The attribution rule: **a purchase is credited to an interaction when the same
client buys within 14 days of it.**

Put `14` in a named cell `AttributionWindow` on the `Config` sheet so the rule
can be tested at 7 or 28 days without rewriting formulas.

### On `Clienteling Activities`, add these columns

```excel
// Column I — attributed_orders
=SUMPRODUCT((tblTransactions[customer_id]=[@customer_id])
           *(tblTransactions[transaction_date]>=[@interaction_date])
           *(tblTransactions[transaction_date]<=[@interaction_date]+AttributionWindow)
           *(tblTransactions[is_return]=FALSE)
           /COUNTIFS(tblTransactions[order_id],tblTransactions[order_id]))

// A1 equivalent, I2 filled down to I3973
=SUMPRODUCT((Transactions!$D$2:$D$15018=B2)
           *(Transactions!$C$2:$C$15018>=E2)
           *(Transactions!$C$2:$C$15018<=E2+$B$2)
           *(Transactions!$P$2:$P$15018=FALSE)
           /COUNTIFS(Transactions!$B$2:$B$15018,Transactions!$B$2:$B$15018))
```

```excel
// Column J — converted_in_window
=IF([@attributed_orders]>0,1,0)

// Column K — attributed_revenue
=SUMIFS(tblTransactions[net_amount],
        tblTransactions[customer_id],[@customer_id],
        tblTransactions[transaction_date],">="&[@interaction_date],
        tblTransactions[transaction_date],"<="&[@interaction_date]+AttributionWindow,
        tblTransactions[is_return],FALSE)

// A1 equivalent, K2 filled down
=SUMIFS(Transactions!$N$2:$N$15018,
        Transactions!$D$2:$D$15018,B2,
        Transactions!$C$2:$C$15018,">="&E2,
        Transactions!$C$2:$C$15018,"<="&E2+$B$2,
        Transactions!$P$2:$P$15018,FALSE)
```

**Read this before quoting the revenue number.** A client with three calls
before one purchase will have that purchase counted against all three
interactions. Conversion *rate* is unaffected — it is measured per interaction
— but summing column K double-counts revenue. For a defensible total, keep only
the last interaction before each order:

```excel
// Column L — is_last_touch
=IF([@attributed_orders]=0, 0,
    IF([@interaction_date]=MAXIFS(tblClienteling[interaction_date],
                                  tblClienteling[customer_id],[@customer_id],
                                  tblClienteling[interaction_date],"<="&
          MINIFS(tblTransactions[transaction_date],
                 tblTransactions[customer_id],[@customer_id],
                 tblTransactions[transaction_date],">="&[@interaction_date],
                 tblTransactions[is_return],FALSE)), 1, 0))
```

Then sum `attributed_revenue` only where `is_last_touch = 1`. This matches
`src/kpis.attribute_clienteling()`, which enforces the same last-touch rule.

### KPI block

```excel
// Appointment conversion rate
=AVERAGEIFS(tblClienteling[converted_in_window],
            tblClienteling[interaction_type],"Appointment")

// Call conversion rate
=AVERAGEIFS(tblClienteling[converted_in_window],
            tblClienteling[interaction_type],"Call")

// Attributed revenue (last touch only)
=SUMIFS(tblClienteling[attributed_revenue],tblClienteling[is_last_touch],1)

// Revenue per interaction
=SUMIFS(tblClienteling[attributed_revenue],tblClienteling[is_last_touch],1)
 / COUNTA(tblClienteling[interaction_id])
```

### Associate performance table

```excel
// Unique associates who logged activity
=SORT(UNIQUE(tblClienteling[associate_id]))

// Associate name
=XLOOKUP([@associate_id],tblAssociates[associate_id],tblAssociates[associate_name],"")

// Interactions logged
=COUNTIFS(tblClienteling[associate_id],[@associate_id])

// Appointments logged
=COUNTIFS(tblClienteling[associate_id],[@associate_id],
          tblClienteling[interaction_type],"Appointment")

// Appointment conversion rate
=AVERAGEIFS(tblClienteling[converted_in_window],
            tblClienteling[associate_id],[@associate_id],
            tblClienteling[interaction_type],"Appointment")

// Attributed revenue
=SUMIFS(tblClienteling[attributed_revenue],
        tblClienteling[associate_id],[@associate_id],
        tblClienteling[is_last_touch],1)
```

Filter the table to associates with a meaningful sample before ranking anyone:

```excel
=FILTER(AssocRange, AssocInteractions>=15,
        "No associate has 15+ logged interactions in this period")
```

Ranking on small denominators produces conversion rates that look dramatic and
mean nothing — an associate with 3 appointments and 2 conversions is not a
67% performer.

---

## Tab 4 — Campaign Effectiveness

`Marketing Campaigns` already carries `attributed_orders`,
`attributed_net_revenue` and `roas` from the pipeline. To rebuild ROAS in Excel
from the transaction data instead:

```excel
// Attributed net revenue for a campaign
=SUMIFS(tblTransactions[net_amount],
        tblTransactions[campaign_id],[@campaign_id],
        tblTransactions[is_return],FALSE)

// ROAS
=SUMIFS(tblTransactions[net_amount],
        tblTransactions[campaign_id],[@campaign_id],
        tblTransactions[is_return],FALSE) / [@spend]

// Blended ROAS across all campaigns
=SUMIFS(tblTransactions[net_amount],tblTransactions[campaign_id],"<>",
        tblTransactions[is_return],FALSE)
 / SUM(tblCampaigns[spend])

// Full-price share of campaign-influenced sales
=SUMIFS(tblTransactions[net_amount],tblTransactions[campaign_id],"<>",
        tblTransactions[is_return],FALSE,tblTransactions[is_full_price],TRUE)
 / SUMIFS(tblTransactions[net_amount],tblTransactions[campaign_id],"<>",
          tblTransactions[is_return],FALSE)

// Full-price share of non-influenced sales
=SUMIFS(tblTransactions[net_amount],tblTransactions[campaign_id],"",
        tblTransactions[is_return],FALSE,tblTransactions[is_full_price],TRUE)
 / SUMIFS(tblTransactions[net_amount],tblTransactions[campaign_id],"",
          tblTransactions[is_return],FALSE)
```

Note the criteria `"<>"` means "not blank" and `""` means "blank". Reversing
them is easy to do and produces two plausible numbers that are exactly wrong.

### ROAS by channel PivotTable

Rows: `channel`. Values: Sum of `spend`, Sum of `attributed_net_revenue`. Then
a calculated field:

```
Insert → PivotTable Analyze → Fields, Items & Sets → Calculated Field
Name:    ROAS
Formula: = attributed_net_revenue / spend
```

Calculated fields operate on the summed values, so this gives
`SUM(revenue)/SUM(spend)` — the correct weighted ROAS, not the average of
per-campaign ratios. Plot as a bar chart sorted descending, and label the bars
with spend so a 30x return on a tiny budget is not mistaken for the biggest
opportunity in the business.

---

## Verification

Before sharing the workbook, check these against
`data/processed/kpi_summary.json`:

| Check | Expected at seed 42 |
|---|---|
| `SUM(net_sales_value)` | 32,194,535 |
| Distinct orders | 9,036 |
| ATV | 3,563 |
| AUR | 2,278 |
| Full-price share | 85.2% |
| Repeat purchase rate | 67.6% |
| Appointment conversion | 53.6% |
| Clients in RFM list | 2,469 |

If Net Sales matches `SUM(net_amount)` exactly, the returns filter is not
working — check that `is_return` imported as a boolean and not as the text
`"TRUE"`, which no comparison against `TRUE` will ever match.
