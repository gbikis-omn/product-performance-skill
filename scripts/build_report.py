#!/usr/bin/env python3
"""
Product Performance Analysis - report builder.

Takes already-pulled raw data (GA4 item-level report, Google Ads spend by
item ID, Meta spend by product ID, optional product-feed field) and builds
one xlsx workbook with:

  - Raw data tabs (for audit/QA, mirrors what a human would VLOOKUP by hand)
  - Master Data tab (everything merged by item_id)
  - Category Analysis tab
  - Brand Analysis tab (optional)
  - Brand & Category Analysis tab (optional, needs Brand Analysis)
  - "<Feed Field> Analysis" tab (optional, one per extra feed field)

All formulas below were reverse-engineered and verified cell-by-cell against
a real client workbook (Kanellopoulos Product Performance, Sept 2026) before
this script was written - see references/mcp_pull_reference.md in this skill
for the verification notes. Do not change the formulas without re-verifying
against a real example.

Usage:
    python3 build_report.py \
        --ga4-csv ga4_items.csv \
        --google-ads-csv google_ads_spend.csv \
        --meta-csv meta_spend.csv \
        --category-levels 3 \
        --include-brand-analysis \
        --output product_performance_report.xlsx \
        [--feed-csv feed.csv --feed-field-name "Custom Label 4"]

Input CSV contracts
--------------------
ga4-csv (required) columns:
    item_id, item_brand,
    item_category, item_category_2, item_category_3, item_category_4, item_category_5,
    items_viewed, items_added_to_cart, items_purchased, item_revenue
  Only the category columns up to --category-levels need to be populated;
  extra columns beyond that are ignored. One row per item_id is expected,
  but the script tolerates (and reconciles) duplicate item_id rows - see
  aggregate_ga4().

google-ads-csv (required) columns: item_id, cost
meta-csv (required) columns: item_id, spend
feed-csv (optional) columns: item_id, <feed-field-name>
"""
import argparse
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

FLOAT_TOL = 0.01  # euro-cents tolerance for reconciliation self-checks

MONEY_FMT = '#,##0.00 "€"'
INT_FMT = '#,##0'
PCT_FMT = '0.00%'

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="E5E7EB")
TOTAL_FONT = Font(bold=True)
GRANDTOTAL_FILL = PatternFill("solid", fgColor="9CA3AF")
GRANDTOTAL_FONT = Font(bold=True, color="FFFFFF")


# --------------------------------------------------------------------------
# Loading & merging
# --------------------------------------------------------------------------

RESERVED_COLUMN_NAMES = {
    "item_id", "item_brand", "Item Brand", "Category", "Meta Spend",
    "Google Ads Spend", "Total Spend", "item_category", "item_category_2",
    "item_category_3", "item_category_4", "item_category_5",
}


def normalize_id_series(s):
    """Make item IDs joinable regardless of how the upstream MCP pull /
    CSV-writing step happened to stringify them. Called on every id column
    from every source (GA4, Google Ads, Meta, feed) so all four end up in
    the same comparable shape before anything gets merged.

    Four real risks this guards against, each one found in practice while
    building/testing this skill - not hypothetical:

      1. Meta's `insights` breakdown by product_id returns a COMPOSITE
         string, e.g. "1013214, FOREL - Μπλούζα ΓΥΝΑΙΚΑ ΜΑΥΡΟ 082.10.01.093"
         (catalog id, comma, product name) - not a clean id. If left as-is
         this NEVER matches GA4/Google Ads item_id, and every row of that
         source's spend silently lands in the "unmatched" bucket with no
         error - the report just quietly shows near-zero spend for that
         channel. Extract the token before the first comma automatically
         whenever a value matches the "<id>, <more text>" shape, for every
         source (harmless no-op on ids that don't look like this).
      2. Whitespace differences ("1234 " vs "1234") - silently breaks an
         exact-match join with no error, just looks like "0 spend".
      3. A numeric ID that went through a float somewhere upstream and
         picked up a trailing ".0" (e.g. pandas/json round-tripping an int
         through a float column) while another source kept it as a clean
         digit string - same silent-join-failure risk.
      4. Different sources report the SAME real id in different letter
         case. Confirmed live on the ProteinMax.gr account (2026-09-14):
         Google Ads' `segments.product_item_id` came back lowercase
         ('im4966') while GA4's item_id / the Merchant Center feed for the
         exact same product is uppercase ('IM4966'). This is a silent,
         total failure, not a partial one - pandas' `merge(on="item_id")`
         is case-sensitive, so it produced ZERO matches out of 1,476 Google
         Ads item_ids that period (97% of which matched once uppercased),
         and 100% of that channel's spend (thousands of euros, both times)
         silently vanished from two already-delivered reports with no
         error - the run even printed a NOTE claiming those items "had ad
         spend but ZERO GA4 activity", which was flatly wrong; they had
         GA4 activity, the join just couldn't see it. Case-fold every id to
         uppercase before comparison so this can never happen silently
         again, for every source, every time - not just when someone
         happens to notice the numbers look low.

    Real alphanumeric IDs/SKUs are left completely untouched by risks 1-3
    (they don't contain ", " followed by more text, and a trailing ".0" is
    only stripped when the rest of the string is purely digits) - so those
    steps never corrupt a legitimate id, they only ever fix the known-bad
    shapes above. Risk 4's uppercase fold is applied unconditionally (safe:
    idempotent on ids that are already uppercase, and it is the join key
    only - see build_master(), the *displayed* item_id in the output comes
    from the GA4 dataframe, which goes through this same normalization, so
    display and join key always match).
    """
    s = s.astype(str).str.strip()

    composite_id = s.str.extract(r"^([^,\s][^,]*),\s+\S", expand=False)
    n_composite = composite_id.notna().sum()
    if n_composite:
        print(f"[build_report] NOTE: auto-parsed {n_composite} composite "
              f"'id, description' value(s) down to just the id (Meta's "
              f"product_id breakdown shape) - if this number looks wrong "
              f"for your source, check the raw pull.", file=sys.stderr)
    s = composite_id.where(composite_id.notna(), s).str.strip()

    is_float_looking = s.str.match(r"^\d+\.0$")
    s = s.where(~is_float_looking, s.str.replace(r"\.0$", "", regex=True))

    upper = s.str.upper()
    n_case_changed = (upper != s).sum()
    if n_case_changed:
        print(f"[build_report] NOTE: case-folded {n_case_changed} id value(s) to "
              f"uppercase for matching (e.g. mixed/lower-case ids like Google Ads' "
              f"segments.product_item_id vs GA4/feed ids that are upper-case) - "
              f"this is expected and required for the join to work; see "
              f"normalize_id_series() docstring risk #4.", file=sys.stderr)
    return upper


def aggregate_ga4(path, category_levels):
    """Load the GA4 item-level pull and collapse to one row per item_id.

    A single item_id can legitimately show up more than once in a raw GA4
    pull when the item_name/category changed mid-period (GA4 keeps history).
    If we naively treat (item_id, item_brand, category...) as the grouping
    key we silently double count that item's views/purchases across the
    two label variants - this exact bug was found and fixed during the
    design of this skill. So: group by item_id ONLY, sum the metrics, and
    pick the brand/category combination with the most views as the
    canonical label for that item_id.
    """
    df = pd.read_csv(path, dtype={"item_id": str})
    if len(df) == 0:
        raise ValueError(
            f"{path} has no data rows (header only). This usually means the "
            f"GA4 pull for the requested period came back empty - per the "
            f"skill's Step 5, that should have been caught and told to the "
            f"user BEFORE reaching this script. Do not proceed - check the "
            f"GA4 query/date range instead of re-running this with empty input."
        )
    df["item_id"] = normalize_id_series(df["item_id"])
    _warn_if_ids_still_look_odd(df["item_id"], "GA4")

    cat_cols = ["item_category", "item_category_2", "item_category_3",
                "item_category_4", "item_category_5"][:category_levels]
    for c in cat_cols:
        if c not in df.columns:
            df[c] = "(not set)"
    if "item_brand" not in df.columns:
        df["item_brand"] = ""
    df["item_brand"] = df["item_brand"].fillna("")

    metric_cols = ["items_viewed", "items_added_to_cart", "items_purchased", "item_revenue"]
    for c in metric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    label_cols = ["item_brand"] + cat_cols
    conflicts = []
    rows = []
    for item_id, g in df.groupby("item_id", sort=False):
        summed = g[metric_cols].sum()
        if g[label_cols].drop_duplicates().shape[0] > 1:
            conflicts.append(item_id)
        canonical = g.loc[g["items_viewed"].idxmax(), label_cols]
        row = {"item_id": item_id, **canonical.to_dict(), **summed.to_dict()}
        rows.append(row)

    out = pd.DataFrame(rows)
    if conflicts:
        print(f"[build_report] NOTE: {len(conflicts)} item_id(s) had more than one "
              f"brand/category label in the GA4 pull (likely renamed mid-period). "
              f"Metrics were summed under item_id; the label with the most views "
              f"was kept. First few: {conflicts[:10]}", file=sys.stderr)

    out["Category"] = out[cat_cols].astype(str).agg(">".join, axis=1)
    return out, conflicts


def _warn_if_ids_still_look_odd(id_series, source_name):
    """normalize_id_series() already auto-parses Meta's composite
    'id, product name' shape, so this is not the primary defense anymore -
    it's a residual sanity net for the case where a comma survived that
    extraction anyway (e.g. a genuinely unexpected id format this skill
    hasn't seen yet). Non-fatal: report it so a human looks, don't guess
    further and don't block the run on a pattern we can't confidently name.
    """
    if len(id_series) == 0:
        return
    comma_frac = id_series.str.contains(",", regex=False).mean()
    if comma_frac > 0:
        example = id_series[id_series.str.contains(",", regex=False)].iloc[0]
        print(f"[build_report] WARNING: {source_name}: {comma_frac:.0%} of item_id "
              f"values still contain a comma after auto-parsing (e.g. {example!r}). "
              f"These will not match GA4/Google Ads item_id and their spend will show "
              f"up as 'unmatched'. This is an id shape the auto-parser didn't "
              f"recognize - look at the raw source data before trusting this report.",
              file=sys.stderr)


def load_spend(path, id_col, value_col, out_col):
    df = pd.read_csv(path, dtype={id_col: str})
    df[id_col] = normalize_id_series(df[id_col])
    _warn_if_ids_still_look_odd(df[id_col], out_col)
    numeric = pd.to_numeric(df[value_col], errors="coerce")
    bad = df[value_col].notna() & numeric.isna()
    if bad.any():
        print(f"[build_report] WARNING: {bad.sum()} row(s) in the {out_col} source had a "
              f"non-numeric value in '{value_col}' (examples: "
              f"{df.loc[bad, value_col].head(5).tolist()}) - treated as 0, not skipped. "
              f"If this is more than a couple of rows, the source column is probably "
              f"wrong (e.g. reading a formatted '€12.34' string instead of a raw "
              f"number) - check before trusting this report's spend totals.",
              file=sys.stderr)
    df[value_col] = numeric.fillna(0.0)
    agg = df.groupby(id_col, as_index=False)[value_col].sum()
    agg = agg.rename(columns={id_col: "item_id", value_col: out_col})
    return agg


def load_feed(path, field_name):
    if field_name in RESERVED_COLUMN_NAMES:
        raise ValueError(
            f"--feed-field-name '{field_name}' collides with a column this script "
            f"already uses internally ({sorted(RESERVED_COLUMN_NAMES)}). Pick a "
            f"different name for the feed field (or rename it in the feed CSV) - "
            f"a silent collision here would overwrite real data."
        )
    df = pd.read_csv(path, dtype={"item_id": str})
    df["item_id"] = normalize_id_series(df["item_id"])
    _warn_if_ids_still_look_odd(df["item_id"], "feed")
    if field_name not in df.columns:
        raise ValueError(f"--feed-field-name '{field_name}' not found in feed CSV columns: {list(df.columns)}")
    df = df[["item_id", field_name]]
    dup_ids = df[df.duplicated("item_id", keep=False)]
    conflicting = dup_ids.groupby("item_id")[field_name].nunique()
    conflicting = conflicting[conflicting > 1]
    if len(conflicting):
        print(f"[build_report] WARNING: {len(conflicting)} item_id(s) in the feed CSV have "
              f"more than one different '{field_name}' value (e.g. a stale + refreshed feed "
              f"pull concatenated together). The first value found is kept; the rest are "
              f"dropped silently otherwise, so this is worth telling the user about. "
              f"First few: {conflicting.index.tolist()[:10]}", file=sys.stderr)
    df = df.drop_duplicates(subset="item_id", keep="first")
    return df


NO_GA4_LABEL = "(no GA4 activity)"


def build_master(ga4_df, google_df, meta_df, feed_df=None, feed_field_name=None):
    """Build the merged master table over the UNION of item_ids seen in any
    of the three sources - not just the ones GA4 happened to report.

    Changed 2026-09-14 (explicit user request, ProteinMax.gr): the original
    version left-joined spend onto ga4_df, so an item_id with real Google
    Ads/Meta spend but zero GA4 views/cart/purchases in the period never
    appeared anywhere in the report - its spend was silently dropped from
    every total. That is a legitimate design choice some users want (spend
    with no matching GA4 row can't be attributed to a real product-level
    conversion story), but it is NOT what this user wants: they want that
    spend visible, grouped under a clearly-labelled bucket, so nothing
    disappears from Total Spend / Grand Total silently.

    Every item_id from GA4 ∪ Google Ads ∪ Meta is now a row in master. Rows
    with no GA4 data get items_viewed/cart/purchased/revenue = 0 and
    item_brand/every item_category* column/Category = NO_GA4_LABEL, so they
    surface as their own explicit "(no GA4 activity)" group in Category
    Analysis, Brand Analysis and Brand & Category Analysis - summable,
    visible, never silently missing - instead of vanishing from the report.
    """
    cat_cols = [c for c in ga4_df.columns if c.startswith("item_category")]

    ga4_ids = set(ga4_df["item_id"])
    all_ids = ga4_ids | set(google_df["item_id"]) | set(meta_df["item_id"])
    spend_only_ids = all_ids - ga4_ids
    n_before = len(all_ids)

    master = pd.DataFrame({"item_id": sorted(all_ids)})
    master = master.merge(ga4_df, on="item_id", how="left")
    assert len(master) == n_before, (
        f"GA4 merge changed row count ({n_before} -> {len(master)}) - "
        f"ga4_df must have a duplicate item_id that survived its own "
        f"groupby, which should be impossible. Do not trust this report."
    )
    metric_cols = ["items_viewed", "items_added_to_cart", "items_purchased", "item_revenue"]
    for c in metric_cols:
        master[c] = master[c].fillna(0.0)
    master["item_brand"] = master["item_brand"].fillna(NO_GA4_LABEL)
    for c in cat_cols:
        master[c] = master[c].fillna(NO_GA4_LABEL)
    master["Category"] = master[cat_cols].astype(str).agg(">".join, axis=1)

    master = master.merge(google_df, on="item_id", how="left")
    assert len(master) == n_before, (
        f"Google Ads merge changed row count ({n_before} -> {len(master)}) - "
        f"google_df must have a duplicate item_id that survived its own "
        f"groupby, which should be impossible. Do not trust this report."
    )
    master = master.merge(meta_df, on="item_id", how="left")
    assert len(master) == n_before, (
        f"Meta merge changed row count ({n_before} -> {len(master)}) - "
        f"meta_df must have a duplicate item_id that survived its own "
        f"groupby, which should be impossible. Do not trust this report."
    )
    master["Google Ads Spend"] = master["Google Ads Spend"].fillna(0.0)
    master["Meta Spend"] = master["Meta Spend"].fillna(0.0)
    master["Total Spend"] = master["Meta Spend"] + master["Google Ads Spend"]

    if feed_df is not None:
        master = master.merge(feed_df, on="item_id", how="left")
        assert len(master) == n_before, (
            f"Feed merge changed row count ({n_before} -> {len(master)})."
        )
        master[feed_field_name] = master[feed_field_name].fillna("(not set)")

    spend_only_google_spend = google_df.loc[
        google_df["item_id"].isin(spend_only_ids), "Google Ads Spend"
    ].sum()
    spend_only_meta_spend = meta_df.loc[
        meta_df["item_id"].isin(spend_only_ids), "Meta Spend"
    ].sum()
    if spend_only_ids:
        print(f"[build_report] NOTE: {len(spend_only_ids)} item_id(s) had ad spend "
              f"(€{spend_only_google_spend:.2f} Google Ads + €{spend_only_meta_spend:.2f} "
              f"Meta) but ZERO GA4 activity in this period. They ARE included in every "
              f"tab (Master Data and every Analysis tab), grouped under brand AND "
              f"category '{NO_GA4_LABEL}', with views/cart/purchases/revenue = 0 - their "
              f"spend is real and counted in every Total Spend / Grand Total, it just "
              f"has no GA4 performance story to attach to. Report this bucket's € amount "
              f"to the user explicitly, same as any other brand/category.",
              file=sys.stderr)

    return master, spend_only_ids, spend_only_google_spend, spend_only_meta_spend


# --------------------------------------------------------------------------
# Ratio helpers
# --------------------------------------------------------------------------

def safe_div(numer, denom):
    return numer / denom if denom else None


def ratio_columns(row):
    views = row["items_viewed"]
    carts = row["items_added_to_cart"]
    purch = row["items_purchased"]
    rev = row["item_revenue"]
    return {
        "AOV": safe_div(rev, purch),
        "Buy To Detail Rate": safe_div(purch, views),
        "Cart To Detail Rate": safe_div(carts, views),
        "Cart To Purchase Rate": safe_div(purch, carts),
    }


# --------------------------------------------------------------------------
# Tab builders (return list-of-dict rows ready to write, plus the grand total
# used for reconciliation)
# --------------------------------------------------------------------------

METRIC_SUM_COLS = ["items_viewed", "items_added_to_cart", "items_purchased",
                    "item_revenue", "Total Spend"]


def build_nested_analysis(master, dim_specs, include_rates):
    """Generalized N-level nested analysis builder. Replaces the old separate
    build_1d_analysis (N=1) and build_brand_category_analysis (N=2, fixed to
    brand+category) with one function that works for any ordered list of
    grouping dimensions - e.g. Category Analysis, Brand Analysis, Brand &
    Category Analysis, and (new) any of those with an extra feed field
    (like Gender) nested in as one more level, all go through this same path.

    dim_specs: ordered list of (df_column_name, display_label) tuples,
    outermost to innermost. Within each level, groups are sorted by revenue
    desc (ties broken by spend desc then label asc). Every level except the
    innermost gets its own "<value> Total" subtotal row after its children;
    the innermost level's rows are plain leaf rows. A final "Grand Total"
    row closes out the whole tab.

    %Spend/%Revenue on every row (leaf or subtotal) are relative to that
    row's OWN immediate parent, not the grand total - e.g. in a 3-level
    Brand>Category>Gender tab, a Gender leaf row's % is of its Category
    subtotal, and that Category subtotal's own % is of its Brand subtotal.
    This generalizes (and is verified byte-identical to) the reference
    workbook's confirmed 2-level rule (brand-category leaf %Spend = row
    spend / that brand's own total spend, confirmed against the live pivot
    formula =G2/$G$63). For N=1 (a plain Category or Brand Analysis) "own
    immediate parent" IS the grand total, since there's only one level -
    same formula, no special-casing needed, and it reproduces the old
    build_1d_analysis's numbers exactly.

    A subtotal row closing level L puts its "<value> Total" text in the
    NEXT column over (dim_specs[L+1]'s label) rather than its own column,
    which is blanked - matching the reference workbook's 2-level layout
    (brand subtotal blank in "Item Brand", text in "Category") generalized
    to however many levels sit below it. Outer levels above L keep their
    real values so a 3+-level subtotal row is still readable in context
    (e.g. a Category-subtotal row still shows which Brand it belongs to).

    include_rates: whether AOV/Buy-To-Detail/Cart-To-Detail/Cart-To-Purchase
    columns are included. Always True when N=1 (matches the always-on
    behavior of the old 1D tabs). For N>=2 this is the caller's choice -
    the reference workbook's own Brand & Category Analysis didn't have
    these (only 5 metrics + the 2 %s), but that was that one pivot's fixed
    config, not a rule this skill has to keep enforcing on every nested
    tab going forward.
    """
    grand_totals = master[METRIC_SUM_COLS].sum()
    labels = [label for _, label in dim_specs]
    want_rates = include_rates or len(dim_specs) == 1

    def row_metrics(m, parent_totals, is_self_total):
        """is_self_total=True for subtotal AND grand-total rows: a subtotal
        row IS the full total for its own group, so %Spend/%Revenue = 100%
        of itself by definition (not "% of its parent" - that number is
        already available by looking at this same group's row one level up,
        or in the matching standalone Category/Brand Analysis tab). This
        matches the reference workbook's verified brand-subtotal behavior
        exactly (=G63/$G$63 = 100%) and generalizes cleanly to any depth:
        only LEAF rows compute a real ratio against their immediate parent.
        """
        d = {
            "Items Viewed": m["items_viewed"],
            "Items Added To Cart": m["items_added_to_cart"],
            "Items Purchased": m["items_purchased"],
            "Item Revenue": m["item_revenue"],
            "Total Spend": m["Total Spend"],
            "% Spend": 1.0 if is_self_total else safe_div(m["Total Spend"], parent_totals["Total Spend"]),
            "% Revenue": 1.0 if is_self_total else safe_div(m["item_revenue"], parent_totals["item_revenue"]),
        }
        if want_rates:
            d.update(ratio_columns(m))
        return d

    def recurse(df_slice, level, prefix, parent_totals):
        col, label = dim_specs[level]
        g = df_slice.groupby(col, as_index=False)[METRIC_SUM_COLS].sum()
        order = g.sort_values(
            by=["item_revenue", "Total Spend", col],
            ascending=[False, False, True],
            kind="mergesort",
        )[col].tolist()
        g_idx = g.set_index(col)

        is_leaf_level = level == len(dim_specs) - 1
        rows = []
        for val in order:
            m = g_idx.loc[val]
            row_labels = dict(prefix)
            row_labels[label] = val
            if is_leaf_level:
                row = {l: row_labels.get(l, "") for l in labels}
                row.update(row_metrics(m, parent_totals, is_self_total=False))
                row["_row_type"] = "leaf"
                rows.append(row)
            else:
                sub_df = df_slice[df_slice[col] == val]
                rows.extend(recurse(sub_df, level + 1, row_labels, m))
                sub_row_labels = dict(prefix)
                sub_row_labels[labels[level + 1]] = f"{val or f'(no {label.lower()})'} Total"
                sub_row = {l: sub_row_labels.get(l, "") for l in labels}
                sub_row.update(row_metrics(m, parent_totals, is_self_total=True))
                sub_row["_row_type"] = "subtotal"
                rows.append(sub_row)
        return rows

    rows = recurse(master, 0, {}, grand_totals)

    # Same "text goes one column in" convention as a subtotal closing level 0
    # (verified against the reference workbook: "Item Brand" blank, "Grand
    # Total" text sits in "Category"). For a single-dimension tab there's no
    # second column, so it goes in the only one there is.
    grand_row = {l: "" for l in labels}
    grand_row[labels[1] if len(labels) > 1 else labels[0]] = "Grand Total"
    grand_row.update(row_metrics(grand_totals, None, is_self_total=True))
    grand_row["_row_type"] = "grandtotal"
    rows.append(grand_row)

    return rows, grand_totals


# --------------------------------------------------------------------------
# xlsx writing
# --------------------------------------------------------------------------

def autosize(ws, ncols, min_width=10, max_width=60):
    for i in range(1, ncols + 1):
        col = get_column_letter(i)
        max_len = min_width
        for cell in ws[col]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)) + 2)
        ws.column_dimensions[col].width = min(max_len, max_width)


MONEY_HEADERS = {"Item Revenue", "Total Spend", "AOV", "item_revenue",
                  "Meta Spend", "Google Ads Spend", "Cost", "Amount Spent"}
PCT_HEADERS = {"% Spend", "% Revenue", "Buy To Detail Rate",
               "Cart To Detail Rate", "Cart To Purchase Rate"}
INT_HEADERS = {"Items Viewed", "Items Added To Cart", "Items Purchased",
               "items_viewed", "items_added_to_cart", "items_purchased"}


def write_raw_tab(wb, title, df):
    ws = wb.create_sheet(title)
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for row in df.itertuples(index=False):
        ws.append(list(row))
    for c_idx, h in enumerate(df.columns, start=1):
        if h in MONEY_HEADERS:
            for cell in ws[get_column_letter(c_idx)][1:]:
                cell.number_format = MONEY_FMT
        elif h in INT_HEADERS:
            for cell in ws[get_column_letter(c_idx)][1:]:
                cell.number_format = INT_FMT
    ws.freeze_panes = "A2"
    autosize(ws, len(df.columns))
    return ws


def write_analysis_tab(wb, title, rows, group_label_col, total_row_predicate):
    """Writes a 1D or 2D analysis tab from a list of row-dicts.
    total_row_predicate(row_dict) -> 'leaf' | 'subtotal' | 'grandtotal'
    """
    ws = wb.create_sheet(title)
    if not rows:
        return ws
    headers = [k for k in rows[0].keys() if not k.startswith("_")]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    for row in rows:
        kind = row.get("_row_type") or total_row_predicate(row)
        ws.append([row.get(h, "") if row.get(h) is not None else "" for h in headers])
        r = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=c_idx)
            if h in MONEY_HEADERS:
                cell.number_format = MONEY_FMT
            elif h in PCT_HEADERS:
                cell.number_format = PCT_FMT
            elif h in INT_HEADERS:
                cell.number_format = INT_FMT
            if kind == "subtotal":
                cell.fill = TOTAL_FILL
                cell.font = TOTAL_FONT
            elif kind == "grandtotal":
                cell.fill = GRANDTOTAL_FILL
                cell.font = GRANDTOTAL_FONT

    ws.freeze_panes = "A2"
    autosize(ws, len(headers))
    return ws


def default_total_predicate(row):
    label = row.get(list(row.keys())[0], "")
    return "grandtotal" if label == "Grand Total" else "leaf"


# --------------------------------------------------------------------------
# Self-check / reconciliation
# --------------------------------------------------------------------------

def reconcile(name, a, b):
    if abs(a - b) > FLOAT_TOL:
        raise AssertionError(
            f"RECONCILIATION FAILED for {name}: {a:.2f} vs {b:.2f} "
            f"(diff {abs(a - b):.2f}). Do NOT trust this report - fix the "
            f"bug before handing it to the user."
        )


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ga4-csv", required=True)
    ap.add_argument("--google-ads-csv", required=True)
    ap.add_argument("--meta-csv", required=True)
    ap.add_argument("--category-levels", type=int, required=True, choices=[1, 2, 3, 4, 5])
    ap.add_argument("--include-brand-analysis", action="store_true")
    ap.add_argument("--feed-csv", default=None)
    ap.add_argument("--feed-field-name", default=None)
    ap.add_argument("--nest-feed-field", action="store_true",
                     help="Also add the feed field as an extra grouping level inside "
                          "Category Analysis, Brand Analysis and Brand & Category "
                          "Analysis (e.g. Category>Gender, Brand>Gender, "
                          "Brand>Category>Gender), on top of its own standalone "
                          "'<field> Analysis' tab. Ask the user explicitly which they "
                          "want - both are legitimate, this is not a default.")
    ap.add_argument("--include-rates-in-nested", action="store_true",
                     help="Include AOV/Buy-To-Detail/Cart-To-Detail/Cart-To-Purchase "
                          "columns in tabs with 2+ grouping levels (Brand & Category "
                          "Analysis, and any tab widened by --nest-feed-field). Off by "
                          "default to match the reference workbook's Brand & Category "
                          "Analysis exactly. Single-dimension tabs (Category Analysis, "
                          "Brand Analysis, a standalone feed-field Analysis) always have "
                          "these regardless of this flag. Ask the user explicitly.")
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    if args.feed_csv and not args.feed_field_name:
        ap.error("--feed-field-name is required when --feed-csv is given")
    if args.nest_feed_field and not args.feed_csv:
        ap.error("--nest-feed-field requires --feed-csv/--feed-field-name")

    ga4_df, conflicts = aggregate_ga4(args.ga4_csv, args.category_levels)
    google_df = load_spend(args.google_ads_csv, "item_id", "cost", "Google Ads Spend")
    meta_df = load_spend(args.meta_csv, "item_id", "spend", "Meta Spend")
    feed_df = load_feed(args.feed_csv, args.feed_field_name) if args.feed_csv else None

    master, spend_only_ids, spend_only_google_spend, spend_only_meta_spend = build_master(
        ga4_df, google_df, meta_df, feed_df, args.feed_field_name
    )

    master_totals = master[METRIC_SUM_COLS].sum()

    wb = Workbook()
    wb.remove(wb.active)

    write_raw_tab(wb, "Google Ads Data", google_df.rename(columns={"item_id": "Item ID", "Google Ads Spend": "Cost"}))
    write_raw_tab(wb, "Meta Data", meta_df.rename(columns={"item_id": "Product ID", "Meta Spend": "Amount Spent"}))
    if feed_df is not None:
        write_raw_tab(wb, "Feed Data", feed_df.rename(columns={"item_id": "Item ID"}))

    master_display_cols = ["item_id", "item_brand"] + \
        [c for c in ga4_df.columns if c.startswith("item_category")] + \
        ["items_viewed", "items_added_to_cart", "items_purchased", "item_revenue",
         "Category", "Meta Spend", "Google Ads Spend", "Total Spend"]
    if feed_df is not None:
        master_display_cols.append(args.feed_field_name)
    write_raw_tab(wb, "Master Data", master[master_display_cols])

    nest_feed = feed_df is not None and args.nest_feed_field
    extra_dim = [(args.feed_field_name, args.feed_field_name)] if nest_feed else []

    def run_tab(title, dims):
        rows, grand_series = build_nested_analysis(master, dims, args.include_rates_in_nested)
        write_analysis_tab(wb, title, rows, dims[0][1], None)
        # Two independent checks, not the same one twice: grand_series is
        # recomputed directly from master (catches a totally broken output),
        # but the leaf-row sum is what actually proves the recursive
        # groupby in build_nested_analysis didn't drop or duplicate any
        # rows on the way down - that's the check that would have caught a
        # real bug (e.g. a NaN grouping key silently excluded by pandas
        # groupby) that comparing two copies of the same master-level sum
        # cannot catch.
        leaf_revenue = sum(r["Item Revenue"] for r in rows if r["_row_type"] == "leaf")
        leaf_spend = sum(r["Total Spend"] for r in rows if r["_row_type"] == "leaf")
        reconcile(f"{title} leaf rows revenue", leaf_revenue, master_totals["item_revenue"])
        reconcile(f"{title} leaf rows spend", leaf_spend, master_totals["Total Spend"])
        reconcile(f"{title} grand total revenue", grand_series["item_revenue"], master_totals["item_revenue"])
        reconcile(f"{title} grand total spend", grand_series["Total Spend"], master_totals["Total Spend"])

    # Category Analysis
    run_tab("Category Analysis", [("Category", "Category")] + extra_dim)

    # Brand Analysis (optional)
    if args.include_brand_analysis:
        run_tab("Brand Analysis", [("item_brand", "Item Brand")] + extra_dim)
        run_tab("Brand & Category Analysis",
                [("item_brand", "Item Brand"), ("Category", "Category")] + extra_dim)

    # Extra feed-field analysis (always its own standalone single-dimension tab,
    # regardless of --nest-feed-field - the two are independent, not mutually exclusive)
    if feed_df is not None:
        run_tab(f"{args.feed_field_name} Analysis", [(args.feed_field_name, args.feed_field_name)])

    wb.save(args.output)

    print(f"[build_report] OK. Wrote {args.output}")
    print(f"[build_report] items in report: {len(master)}")
    print(f"[build_report] grand total: views={master_totals['items_viewed']:.0f} "
          f"cart={master_totals['items_added_to_cart']:.0f} "
          f"purchases={master_totals['items_purchased']:.0f} "
          f"revenue={master_totals['item_revenue']:.2f} "
          f"spend={master_totals['Total Spend']:.2f}")
    if conflicts:
        print(f"[build_report] WARNING: {len(conflicts)} item_id(s) had inconsistent "
              f"brand/category labels across the GA4 pull - tell the user.")
    if spend_only_ids:
        print(f"[build_report] WARNING: {len(spend_only_ids)} item(s) had ad spend "
              f"(€{spend_only_google_spend:.2f} Google Ads + €{spend_only_meta_spend:.2f} "
              f"Meta) but no GA4 activity - INCLUDED in every tab under the "
              f"'{NO_GA4_LABEL}' brand/category bucket (not excluded). Tell the user "
              f"this bucket's € amount, same as any other brand/category.")


if __name__ == "__main__":
    main()
