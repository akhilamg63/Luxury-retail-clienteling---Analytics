// =========================================================================
// Luxury Retail CRM & Clienteling Analytics - Power Query (M) setup
//
// Paste each block below into its own blank query in Power BI Desktop
// (Home > Transform data > New Source > Blank Query > Advanced Editor).
//
// Set the folder path ONCE in the SourceFolder query, then every other
// query reads from it. Do not hardcode a path into the individual queries -
// changing machine then means editing nine queries instead of one.
//
// After loading, build the relationships described in star_schema.md and
// add the measures from measures.dax.
// =========================================================================


// -------------------------------------------------------------------------
// Query name: SourceFolder    (this is a PARAMETER, not a table)
// -------------------------------------------------------------------------
// Home > Manage Parameters > New Parameter
//   Name: SourceFolder, Type: Text
//   Current value: the full path to data/processed in your clone, e.g.
//   C:\Users\you\luxury-retail-clienteling-analytics\data\processed
//
// If you would rather use a plain query than a parameter:
//
//   "C:\path\to\luxury-retail-clienteling-analytics\data\processed"


// -------------------------------------------------------------------------
// Query name: Transactions          (fact table)
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\transactions.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    // is_return and is_full_price MUST land as logical, not text. If they
    // import as text, every returns filter silently matches nothing and Net
    // Sales quietly includes returns - a failure that produces plausible
    // numbers, which is what makes it dangerous.
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"transaction_id", type text},
            {"order_id", type text},
            {"transaction_date", type date},
            {"customer_id", type text},
            {"sku", type text},
            {"boutique_id", type text},
            {"associate_id", type text},
            {"channel", type text},
            {"quantity", Int64.Type},
            {"unit_list_price", type number},
            {"discount_pct", type number},
            {"unit_net_price", type number},
            {"gross_amount", type number},
            {"net_amount", type number},
            {"unit_cost", type number},
            {"is_return", type logical},
            {"is_full_price", type logical},
            {"interaction_id", type text},
            {"campaign_id", type text}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Customers
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\customers.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"customer_id", type text},
            {"client_name", type text},
            {"join_date", type date},
            {"acquisition_source", type text},
            {"preferred_boutique_id", type text},
            {"consent_status", type text},
            {"client_tier", type text},
            {"country", type text},
            {"region", type text},
            {"email_opt_in", type logical},
            {"churn_date", type date},
            {"primary_category_affinity", type text}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Products
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\products.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"sku", type text},
            {"product_name", type text},
            {"category", type text},
            {"subcategory", type text},
            {"collection", type text},
            {"season", type text},
            {"list_price", type number},
            {"unit_cost", type number},
            {"gross_margin_pct", type number}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Boutiques
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\boutiques.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"boutique_id", type text},
            {"boutique_name", type text},
            {"city", type text},
            {"country", type text},
            {"region", type text},
            {"boutique_format", type text},
            {"opening_date", type date},
            {"traffic_weight", type number},
            {"annual_sales_target", type number}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Associates
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\associates.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"associate_id", type text},
            {"associate_name", type text},
            {"boutique_id", type text},
            {"role", type text},
            {"tenure_years", type number},
            {"performance_index", type number}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Clienteling
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\clienteling_interactions.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"interaction_id", type text},
            {"customer_id", type text},
            {"associate_id", type text},
            {"boutique_id", type text},
            {"interaction_date", type date},
            {"interaction_type", type text},
            {"converted", type logical},
            {"outcome", type text}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Campaigns
// -------------------------------------------------------------------------
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\campaigns.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"campaign_id", type text},
            {"campaign_name", type text},
            {"audience_segment", type text},
            {"channel", type text},
            {"start_date", type date},
            {"end_date", type date},
            {"spend", type number},
            {"targeted_customers", Int64.Type},
            {"impressions", Int64.Type},
            {"clicks", Int64.Type},
            {"attributed_orders", Int64.Type},
            {"attributed_customers", Int64.Type},
            {"attributed_net_revenue", type number},
            {"roas", type number}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: RFM Segments        (optional - the Python segmentation)
// -------------------------------------------------------------------------
// Load this if you want Power BI to use the pipeline's segments rather than
// recomputing them in DAX. Useful for a report that must agree with the
// Streamlit dashboard figure for figure.
let
    Source = Csv.Document(
        File.Contents(SourceFolder & "\rfm_segments.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"customer_id", type text},
            {"recency_days", Int64.Type},
            {"frequency", Int64.Type},
            {"monetary", type number},
            {"r_score", Int64.Type},
            {"f_score", Int64.Type},
            {"m_score", Int64.Type},
            {"rfm_score", Int64.Type},
            {"segment", type text}
        }
    )
in
    Typed


// -------------------------------------------------------------------------
// Query name: Date            (generated, not loaded from file)
// -------------------------------------------------------------------------
// Mark this as the date table after loading: Table tools > Mark as date
// table > Date. Without that, SAMEPERIODLASTYEAR returns wrong results at
// year boundaries.
let
    MinDate = List.Min(Transactions[transaction_date]),
    MaxDate = List.Max(Transactions[transaction_date]),
    StartDate = #date(Date.Year(MinDate), 1, 1),
    EndDate = #date(Date.Year(MaxDate), 12, 31),
    DayCount = Duration.Days(EndDate - StartDate) + 1,
    Dates = List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),
    AsTable = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),
    Typed = Table.TransformColumnTypes(AsTable, {{"Date", type date}}),
    WithYear = Table.AddColumn(Typed, "Year", each Date.Year([Date]), Int64.Type),
    WithMonthNo = Table.AddColumn(WithYear, "Month Number", each Date.Month([Date]), Int64.Type),
    WithMonth = Table.AddColumn(WithMonthNo, "Month", each Date.ToText([Date], "MMM"), type text),
    WithYearMonth = Table.AddColumn(WithMonth, "Year Month", each Date.ToText([Date], "yyyy-MM"), type text),
    WithQuarter = Table.AddColumn(WithYearMonth, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    WithDow = Table.AddColumn(WithQuarter, "Day Of Week", each Date.ToText([Date], "ddd"), type text),
    WithWeekend = Table.AddColumn(WithDow, "Is Weekend", each Date.DayOfWeek([Date], Day.Monday) > 4, type logical),
    WithSeason = Table.AddColumn(
        WithWeekend, "Retail Season",
        each if Date.Month([Date]) >= 3 and Date.Month([Date]) <= 8 then "SS" else "FW",
        type text
    )
in
    WithSeason


// -------------------------------------------------------------------------
// Query name: Attribution Parameter    (disconnected)
// -------------------------------------------------------------------------
// Drives the clienteling attribution window slicer. It has NO relationships
// by design - the measures read it with SELECTEDVALUE.
let
    Source = {7, 14, 21, 28},
    AsTable = Table.FromList(Source, Splitter.SplitByNothing(), {"Days"}),
    Typed = Table.TransformColumnTypes(AsTable, {{"Days", Int64.Type}})
in
    Typed


// =========================================================================
// After loading
// =========================================================================
// 1. Build the relationships in star_schema.md. Keep them single-direction.
//    Set the Clienteling -> Boutiques and Clienteling -> Date relationships
//    INACTIVE, or Power BI will reject the model as ambiguous.
// 2. Mark Date as the date table.
// 3. Sort Month by Month Number, or months render alphabetically.
// 4. Paste the measures from measures.dax.
// 5. Verify before building visuals: [Net Sales] must be LOWER than
//    SUM(Transactions[net_amount]) over all rows. If they are equal, the
//    is_return column imported as text and the returns filter is doing
//    nothing.
// =========================================================================
