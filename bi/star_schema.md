# Star schema modelling guide

How to wire the seven generated tables into a star schema in Power BI or
Tableau. The target is one fact table, conformed dimensions, and every
relationship single-direction.

## Model diagram

```
                      ┌──────────────────┐
                      │      Date        │
                      │  (generated)     │
                      │  PK: Date        │
                      └────────┬─────────┘
                               │ 1
                               │
                               │ *
┌───────────────┐  1     *  ┌──┴───────────────────────┐  *     1  ┌───────────────┐
│   Customers   ├──────────►│                          │◄──────────┤   Products    │
│ PK customer_id│           │       Transactions       │           │   PK sku      │
└───────────────┘           │          (fact)          │           └───────────────┘
                            │                          │
┌───────────────┐  1     *  │  FK customer_id          │  *     1  ┌───────────────┐
│   Boutiques   ├──────────►│  FK sku                  │◄──────────┤  Associates   │
│PK boutique_id │           │  FK boutique_id          │           │PK associate_id│
└───────────────┘           │  FK associate_id         │           └───────────────┘
                            │  FK campaign_id          │
┌───────────────┐  1     *  │  FK interaction_id       │
│   Campaigns   ├──────────►│  transaction_date        │
│PK campaign_id │           │  quantity, net_amount    │
└───────────────┘           │  is_return, is_full_price│
                            └──────────────────────────┘
                                      ▲ *
                                      │
                                      │ 1
                            ┌─────────┴──────────┐
                            │    Clienteling     │
                            │ PK interaction_id  │
                            │ FK customer_id     │
                            │ FK associate_id    │
                            └────────────────────┘
```

## Tables and roles

| Table | Role | Grain | Key |
|---|---|---|---|
| `transactions` | Fact | One row per order line | `transaction_id` |
| `customers` | Dimension | One row per client | `customer_id` |
| `products` | Dimension | One row per SKU | `sku` |
| `boutiques` | Dimension | One row per boutique | `boutique_id` |
| `associates` | Dimension | One row per sales associate | `associate_id` |
| `campaigns` | Dimension | One row per campaign | `campaign_id` |
| `clienteling_interactions` | Fact (secondary) | One row per interaction | `interaction_id` |
| `Date` | Dimension | One row per day | `Date` |

## Relationships to create

| From | To | Cardinality | Direction | Active |
|---|---|---|---|---|
| `transactions[customer_id]` | `customers[customer_id]` | Many-to-one | Single | Yes |
| `transactions[sku]` | `products[sku]` | Many-to-one | Single | Yes |
| `transactions[boutique_id]` | `boutiques[boutique_id]` | Many-to-one | Single | Yes |
| `transactions[associate_id]` | `associates[associate_id]` | Many-to-one | Single | Yes |
| `transactions[campaign_id]` | `campaigns[campaign_id]` | Many-to-one | Single | Yes |
| `transactions[transaction_date]` | `Date[Date]` | Many-to-one | Single | Yes |
| `clienteling_interactions[customer_id]` | `customers[customer_id]` | Many-to-one | Single | Yes |
| `clienteling_interactions[associate_id]` | `associates[associate_id]` | Many-to-one | Single | Yes |
| `clienteling_interactions[boutique_id]` | `boutiques[boutique_id]` | Many-to-one | Single | **No** |
| `clienteling_interactions[interaction_date]` | `Date[Date]` | Many-to-one | Single | **No** |

### Why two relationships are inactive

`clienteling_interactions` already reaches `Boutiques` through `Customers`
(`preferred_boutique_id`) and through `Associates`. Adding a third active path
creates an ambiguous filter path, and Power BI will refuse it. Keep the direct
boutique relationship inactive and activate it only where needed:

```
Interactions By Boutique =
CALCULATE (
    COUNTROWS ( Clienteling ),
    USERELATIONSHIP ( Clienteling[boutique_id], Boutiques[boutique_id] )
)
```

The date relationship is inactive for the same reason: both facts cannot filter
the same date dimension actively without creating an ambiguous model. Activate
it per measure:

```
Interactions In Period =
CALCULATE (
    COUNTROWS ( Clienteling ),
    USERELATIONSHIP ( Clienteling[interaction_date], 'Date'[Date] )
)
```

## Date table

Create it in DAX and mark it as a date table (Table tools → Mark as date
table). Without this, time intelligence functions such as `SAMEPERIODLASTYEAR`
return wrong results at year boundaries.

```
Date =
VAR MinDate = MIN ( Transactions[transaction_date] )
VAR MaxDate = MAX ( Transactions[transaction_date] )
RETURN
ADDCOLUMNS (
    CALENDAR ( DATE ( YEAR ( MinDate ), 1, 1 ), DATE ( YEAR ( MaxDate ), 12, 31 ) ),
    "Year",          YEAR ( [Date] ),
    "Month Number",  MONTH ( [Date] ),
    "Month",         FORMAT ( [Date], "MMM" ),
    "Year Month",    FORMAT ( [Date], "YYYY-MM" ),
    "Quarter",       "Q" & QUARTER ( [Date] ),
    "Day Of Week",   FORMAT ( [Date], "ddd" ),
    "Is Weekend",    WEEKDAY ( [Date], 2 ) > 5,
    "Retail Season", IF ( MONTH ( [Date] ) >= 3 && MONTH ( [Date] ) <= 8, "SS", "FW" )
)
```

Sort `Month` by `Month Number`, otherwise months render alphabetically.

## Attribution parameter table

Disconnected, single-column table that drives the clienteling window. It has no
relationships by design — it is read with `SELECTEDVALUE`, not used as a filter.

```
Attribution Parameter =
GENERATESERIES ( 7, 30, 7 )
```

Rename the column to `Days`, then expose it as a slicer. The DAX measures read
it through `[Attribution Window Days]`, so changing the slicer re-runs
conversion and attributed revenue at 7, 14, 21 or 28 days.

## Modelling rules worth enforcing

**Keep relationships single-direction.** Bi-directional filtering is tempting
for the clienteling table, but it introduces ambiguity as soon as a second fact
is added and makes measures behave differently depending on what is on the
canvas. Use `CROSSFILTER` in the one measure that needs it instead.

**Hide every key column from report view.** Users should never drag
`customer_id` onto a canvas; they should use `client_name` or a segment. Hiding
keys also prevents accidental `COUNT(customer_id)` measures that look right and
are not.

**Never join Clienteling to Transactions directly.** The relationship is
many-to-many through `customer_id`, and a physical join fans out transaction
rows — one order preceded by three calls becomes three rows and triples
revenue. Keep them as separate facts sharing the `Customers` and `Associates`
dimensions, and resolve attribution in a measure (see `measures.dax` §7) or
upstream in the pipeline.

**Set `is_return` and `is_full_price` to TRUE/FALSE, not text.** The DAX
measures compare against `FALSE()`. If the CSV import types them as text,
every return filter silently matches nothing and Net Sales will include
returns — a failure mode that produces plausible-looking numbers, which is what
makes it dangerous. Verify after import: `Net Sales` must be lower than
`SUM(net_amount)` over all rows.

**Mark the fact table's numeric columns as non-summarisable where they are
rates.** `discount_pct` and `gross_margin_pct` summed across rows are
meaningless; set their default aggregation to Average or None.
