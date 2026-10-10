# Tableau calculated fields

Tableau equivalents of every measure in `measures.dax`. The arithmetic is
identical; what changes is how each tool handles grain.

The important structural difference: DAX resolves "distinct orders" with
`DISTINCTCOUNT` inside a measure, whereas Tableau needs either `COUNTD` or a
Level of Detail (LOD) expression to escape the grain of the current view. Where
an LOD is required, it is marked and explained — getting this wrong is the
usual reason a Tableau ATV disagrees with the same figure in Power BI.

## Data source setup

Transactions is the primary table. Left-join `Products` on `sku`, `Boutiques`
on `boutique_id`, `Associates` on `associate_id` and `Customers` on
`customer_id`. Keep `Clienteling Activities` and `Marketing Campaigns` as
**separate data sources** and relate them with a blend or a relationship on
`customer_id` / `campaign_id` — joining them to the fact table fans out the
transaction rows and silently doubles revenue.

---

## 1. Net Sales (excluding returns)

```
// Net Sales
IF [Is Return] = FALSE THEN [Net Amount] END
```

Used as `SUM([Net Sales])`. The `IF` with no `ELSE` returns Null for returned
lines, and `SUM` ignores Nulls — cleaner than `ELSE 0`, because it also keeps
returned lines out of `COUNT` and `AVG` if anyone reuses the field.

```
// Gross Sales
IF [Is Return] = FALSE THEN [Gross Amount] END

// Returned Sales
IF [Is Return] = TRUE THEN [Net Amount] END

// Return Rate %
SUM([Returned Sales]) / (SUM([Net Sales]) + SUM([Returned Sales]))
```

## 2. Average Transaction Value (ATV)

```
// Transaction Count
COUNTD(IF [Is Return] = FALSE THEN [Order Id] END)

// ATV
SUM([Net Sales]) / [Transaction Count]
```

Order-level, not line-level. If the view is already at order grain, the simpler
`{FIXED [Order Id] : SUM([Net Sales])}` averaged across orders gives the same
answer — but `COUNTD` is correct at every grain, so prefer it.

## 3. Average Unit Retail (AUR)

```
// Units Sold
IF [Is Return] = FALSE THEN [Quantity] END

// AUR
SUM([Net Sales]) / SUM([Units Sold])

// UPT (units per transaction)
SUM([Units Sold]) / [Transaction Count]
```

`ATV = AUR × UPT`. Showing all three lets a merchandising team tell "bigger
baskets" apart from "more expensive pieces".

## 4. Repeat Purchase Rate

This one needs an LOD. The orders-per-client count has to be computed at client
grain regardless of what is on the view.

```
// Orders Per Client  (LOD - fixed at client grain)
{FIXED [Customer Id] : COUNTD(IF [Is Return] = FALSE THEN [Order Id] END)}

// Is Repeat Client
[Orders Per Client] >= 2

// Repeat Purchase Rate %
COUNTD(IF [Is Repeat Client] THEN [Customer Id] END)
/
COUNTD(IF [Orders Per Client] >= 1 THEN [Customer Id] END)
```

The denominator counts clients with at least one purchase, not every client in
the CRM, so prospect records cannot drag the rate down.

Caveat worth knowing: `FIXED` ignores dimension filters but respects context
filters. If the dashboard has a date filter, promote it to a context filter, or
`Orders Per Client` will count a client's entire history while the numerator
respects the date range — producing a rate that drifts as users change dates.

## 5. Full-Price Sales Share

```
// Full Price Sales
IF [Is Return] = FALSE AND [Is Full Price] = TRUE THEN [Net Amount] END

// Full Price Share %
SUM([Full Price Sales]) / SUM([Net Sales])

// Markdown Depth %
1 - (SUM([Net Sales]) / SUM([Gross Sales]))
```

By value, not by line count.

## 6. Campaign ROAS

Campaign spend lives at campaign grain. Summing it against transaction rows
counts the same budget once per attributed order.

```
// Campaign Spend  (from the Campaigns source, or LOD if blended in)
{FIXED [Campaign Id] : MAX([Spend])}

// Campaign Attributed Revenue
IF [Is Return] = FALSE AND NOT ISNULL([Campaign Id]) THEN [Net Amount] END

// Campaign ROAS
SUM([Campaign Attributed Revenue]) / SUM([Campaign Spend])

// Cost Per Attributed Order
SUM([Campaign Spend])
/ COUNTD(IF [Is Return] = FALSE AND NOT ISNULL([Campaign Id]) THEN [Order Id] END)
```

## 7. Clienteling Conversion Rate

Tableau cannot express "any order within 14 days" in a single calculated field
the way DAX can, because there is no row-context iterator over a second table.
Two options:

**Option A — relationship with a calculation (preferred).**
Relate `Clienteling Activities` to `Transactions` on `Customer Id`, then:

```
// Days To Purchase
DATEDIFF('day', [Interaction Date], [Transaction Date])

// In Attribution Window
[Days To Purchase] >= 0 AND [Days To Purchase] <= [Attribution Window]

// Converted Interaction  (LOD at interaction grain)
{FIXED [Interaction Id] : MAX(IF [In Attribution Window] THEN 1 ELSE 0 END)}

// Clienteling Conversion Rate %
AVG([Converted Interaction])
```

`[Attribution Window]` is a parameter (integer, default 14) so the window can be
tested in the view without republishing.

**Option B — resolve upstream.** Materialise the attribution in the pipeline and
expose a boolean `converted_in_window` column. This is what
`src/kpis.attribute_clienteling()` already does, and it is the faster option on
large extracts. Option A is only worth it when the business genuinely wants to
drag the window parameter around.

```
// Appointment Conversion Rate %
AVG(IF [Interaction Type] = 'Appointment' THEN [Converted Interaction] END)

// Clienteling Attributed Revenue
SUM(IF [In Attribution Window] AND [Is Return] = FALSE THEN [Net Amount] END)

// Revenue Per Interaction
[Clienteling Attributed Revenue] / COUNTD([Interaction Id])
```

One caution on Option A: because a client can have several interactions before
one purchase, the same order can satisfy the window for more than one
interaction. Conversion *rate* is unaffected (it is measured per interaction),
but attributed *revenue* will overstate unless last-touch is enforced:

```
// Last Touch Interaction Id  (LOD at order grain)
{FIXED [Order Id] : MAX(IF [In Attribution Window] THEN [Interaction Date] END)}

// Is Last Touch
[Interaction Date] = [Last Touch Interaction Id]
```

Then filter attributed revenue on `[Is Last Touch]`.

## 8. Executive supporting fields

```
// Net Sales LY
LOOKUP(SUM([Net Sales]), -12)        // on a monthly axis

// Net Sales YoY %
(SUM([Net Sales]) - LOOKUP(SUM([Net Sales]), -12)) / LOOKUP(SUM([Net Sales]), -12)

// Target Achievement %
SUM([Net Sales]) / SUM([Annual Sales Target])

// Client Spend  (LOD, for concentration analysis)
{FIXED [Customer Id] : SUM([Net Sales])}

// Client Spend Rank
RANK(SUM([Net Sales]), 'desc')
```

For "share of revenue from the top 10% of clients", put `Customer Id` on detail,
`Client Spend Rank` as a table calculation, and use a `Top N` parameter — or
precompute the decile in the pipeline, which keeps the workbook responsive on
larger client books.

## Formatting

| Field | Format |
|---|---|
| Net Sales, Gross Sales, Campaign Spend | Currency, 0 dp, thousands separator |
| ATV, AUR | Currency, 0 dp |
| Repeat Purchase Rate %, Full Price Share %, conversion rates | Percentage, 1 dp |
| ROAS | Number, 2 dp, suffix `x` |
| UPT | Number, 2 dp |
