# MCP pull reference — verified calls

Every field/resource name below was confirmed against the live tool metadata
(not assumed) while this skill was built — see the session that authored it.
If a pull fails with "field not found" or similar, the account's API version
may differ; re-verify with the metadata/search tools mentioned in each
section rather than guessing an alternate field name.

## 1. GA4 — item-level performance (`ga4-mcp`)

Tool: `ga4_run_report`

```json
{
  "property_id": "<GA4 property id, digits only>",
  "dimensions": ["itemId", "itemBrand", "itemCategory", "itemCategory2", "itemCategory3"],
  "metrics": ["itemsViewed", "itemsAddedToCart", "itemsPurchased", "itemRevenue"],
  "start_date": "<YYYY-MM-DD or NdaysAgo>",
  "end_date": "<YYYY-MM-DD or 'yesterday'>",
  "limit": 100000,
  "order_bys": [{"metric": {"metric_name": "itemsViewed"}, "desc": true}]
}
```

- Only include as many `itemCategoryN` dimensions as the user asked for (1-5 levels). `itemCategory` = level 1, `itemCategory2`..`itemCategory5` = levels 2-5.
- Metric names are confirmed correct: `itemsViewed`, `itemsAddedToCart` (NOT `addToCarts` — that is session-scoped and errors on an item-level report), `itemsPurchased`, `itemRevenue`. Confirmed via `ga4_get_metadata` before use.
- Response can be large (thousands of rows) and blow the tool's token limit — the tool auto-saves it as a receipt and you retrieve it via `get_tool_receipt(receipt_id, offset_bytes=...)`, base64-encoded, in ~390KB chunks with a `next_offset_bytes` field.
  **Always pass forward the exact `next_offset_bytes` value the previous chunk returned — never compute or guess an offset yourself.** Confirmed live (2026-09-11): guessing an offset that was off by only 3 bytes silently produced a misaligned chunk (the call still succeeded, no error) that would have corrupted the reassembled JSON if concatenated. After collecting all chunks in order, decode+concatenate the base64 `data` fields and **verify the result's sha256 against the `sha256` field in the last chunk** before trusting it (`hashlib.sha256(full_bytes).hexdigest() == chunk['sha256']`) — this is the only way to be sure nothing was skipped or duplicated.
- **A single item_id can appear on more than one row** if its brand/category changed mid-period (GA4 keeps the label that was active on each day). Do not treat (item_id, brand, category) as the unique key — `scripts/build_report.py` already handles this (sums metrics by item_id, keeps the label with the most views), but the raw CSV you hand it should still be at item_id+label granularity (don't pre-aggregate yourself).
- Convert the pulled rows to a CSV with columns: `item_id, item_brand, item_category, item_category_2, item_category_3, item_category_4, item_category_5, items_viewed, items_added_to_cart, items_purchased, item_revenue` (only the category columns the user asked for need real values — leave the rest out or blank, the script handles up to `--category-levels`).

## 2. Google Ads — spend per product/item ID (`gads-mcp`)

Tool: `ads_search` (or `ads_export_report` for very large accounts)

```
SELECT segments.product_item_id, metrics.cost_micros
FROM shopping_performance_view
WHERE segments.date DURING <LAST_7_DAYS | LAST_30_DAYS | BETWEEN 'YYYY-MM-DD' AND 'YYYY-MM-DD'>
```

- Verified via `ads_get_field_metadata(resource_name="segments")`: `segments.product_item_id` is a real STRING field, and its `selectable_with` list explicitly includes `shopping_performance_view` and `metrics.cost_micros`. Do not substitute a different resource/field name without re-checking that metadata call.
- `metrics.cost_micros` is in micros — divide by 1,000,000 to get the real spend before writing the CSV.
- `shopping_performance_view` only returns rows for Shopping / Performance Max campaigns with product-level data. If the account has none, the query legitimately returns 0 rows — tell the user rather than assuming a bug.
- Sum `cost` per `item_id` if the same item_id appears on multiple rows (e.g. split by device/network segment you didn't select) before writing the CSV — or just let `scripts/build_report.py` do it (it groups and sums by item_id on load).
- Output CSV columns: `item_id, cost`.

## 3. Meta Ads — spend per product ID (`meta-mcp`)

**Use `meta_export_insights` as the primary tool, not `meta_read`.** `meta_read` on this
resource caps at 200 rows per call with no reliable cursor (see pagination note below) —
for any account with more than ~200 spending products (most real accounts) you will burn
dozens of round trips and still risk gaps. `meta_export_insights` returns the full result
set in one durable CSV artifact, verified live on an account with 4,336 rows.

```json
{
  "ad_account_id": "<numeric ad account id, no act_ prefix>",
  "params": {
    "breakdowns": ["product_id"],
    "fields": ["spend"],
    "time_range": {"since": "YYYY-MM-DD", "until": "YYYY-MM-DD"}
  },
  "format": "csv",
  "max_rows": 5000
}
```

- **Do NOT pass `"level": "ad"`.** Confirmed live (2026-09-11): adding `"level": "ad"` to
  this exact call makes `meta_export_insights` fail with a generic error. Removing it
  (leave level unset) made the identical query succeed. This is not documented anywhere —
  just don't pass it for this specific tool+breakdown combination.
- If `meta_export_insights` still fails after that, only then fall back to `meta_read`
  with `params.limit` (200/page) and treat its `after` param as an **unverified,
  undocumented** offset-in-base64 (e.g. `after: base64("200")` returned page 2 in one
  live test) — it is NOT a real Graph API cursor, has no confirmed gap/duplicate
  guarantees, and should be a last resort, not the default path.
- `product_id` is a **documented, real Meta breakdown** for catalog/Dynamic Product Ads
  reporting (confirmed via Meta's own developer docs during this skill's design — it is
  explicitly called out as a "high-cardinality breakdown").
- **The returned `product_id` field is a COMPOSITE string, not a clean ID** — confirmed
  live, e.g. `"1013214, FOREL - Μπλούζα ΓΥΝΑΙΚΑ ΜΑΥΡΟ 082.10.01.093"` (catalog id, comma,
  product name). **You do NOT need to parse this yourself** — `scripts/build_report.py`'s
  `normalize_id_series()` auto-detects and strips this shape on every id column it loads
  (GA4/Google Ads/Meta/feed alike), verified live to produce byte-identical output whether
  you hand it the raw composite string or a pre-parsed clean id. Just write Meta's
  `product_id` value straight into the CSV's `item_id` column as-is — the script handles
  it. (If you're writing the CSV in some other tool that isn't this script, then yes,
  parse it yourself: split on the first comma, keep the part before it.)
- High-cardinality breakdowns can be unreliable over long date ranges in one call. If a
  pull for the full requested period comes back empty, truncated, or errors even after
  removing `level`, **split the range into daily calls and sum**, rather than silently
  reporting partial/zero spend as if it were complete.
- This breakdown only returns data for campaigns actually running Dynamic/Catalog ads. An
  account running only non-catalog campaigns will legitimately return 0 rows — tell the
  user.
- **Important cross-platform gotcha to flag to the user**: the id half of Meta's
  `product_id` is whatever value is in the product's Meta catalog feed (usually
  `retailer_id`/`content_id`). It only lines up with GA4's `item_id` / Google's
  `product_item_id` if the client's Meta catalog feed and Google Merchant Center feed use
  the same ID scheme (this is normal but not guaranteed — confirm with the user if match
  rates look low after the merge).
- Sum `spend` per item_id if the same id appears on multiple rows before writing the CSV
  (or let the script do it).
- Output CSV columns: `item_id, spend` (the parsed id, not the raw composite string).

## 4. Optional product feed field (`mc-mcp` or a feed URL)

If the user wants an extra dimension from their product feed (mirrors the demo workbook's "Custom Label 4"):

- If they have Merchant Center connected via `mc-mcp`: per this account's own past-verified gotchas, the `products` read has no working pagination and no functional filter param — the reliable path is to fetch the **raw feed XML/CSV URL** directly (`data_sources` resource gives you that URL) and parse the desired field (e.g. `custom_label_4`, `product_type`, `google_product_category`) yourself from the raw feed rather than trusting the MCP's `products` tool for full coverage.
- If they just hand you a feed file/URL, download and parse it directly (id column = `id`/`g:id`, whatever field they asked for).
- Output CSV columns: `item_id, <field name the user asked for>` — pass that exact field name to `--feed-field-name` on the script.

## 5. Date range translation

Whatever preset the user picks, resolve it to concrete dates *before* calling any tool, and use the **same** resolved dates across all three pulls (GA4/Google Ads/Meta) — a mismatch here silently produces an apples-to-oranges report.

| User says | GA4 `start_date`/`end_date` | Google Ads `segments.date DURING` | Meta `time_range` |
|---|---|---|---|
| Τελευταίες 7 μέρες | `7daysAgo` / `yesterday` | `LAST_7_DAYS` | since/until = today-7 .. yesterday |
| Τελευταίες 14 μέρες | `14daysAgo` / `yesterday` | `LAST_14_DAYS` | since/until = today-14 .. yesterday |
| Τελευταίες 30 μέρες | `30daysAgo` / `yesterday` | `LAST_30_DAYS` | since/until = today-30 .. yesterday |
| 1η του μήνα → χθες | explicit `YYYY-MM-01` / yesterday's date | `BETWEEN 'YYYY-MM-01' AND 'YYYY-MM-DD'` | since/until explicit |
| Custom | explicit dates | `BETWEEN ... AND ...` | since/until explicit |
