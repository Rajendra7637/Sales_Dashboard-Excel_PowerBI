# Sales Data MIS Dashboard

Excel-based MIS workbook analyzing 500 e-commerce orders (Jan 2023–Dec 2024, $79,227 in total revenue) to identify concentration risk, order-fulfillment leakage, and data quality gaps that a surface-level pivot table misses.

## Key Findings

**1. Revenue is dangerously concentrated in one product and two categories.**
Electronics and Furniture make up 87.2% of total revenue ($39,218 and $29,894 respectively), and a single SKU — the Standing Desk — accounts for **29.2% of all revenue on its own** ($23,174 across 500 orders). Losing that one product line, a competitor undercutting it, or a supply disruption would hit revenue harder than any single regional or channel shift in the dataset.

**2. Nearly a quarter of orders never convert into revenue.**
124 of 500 orders (24.8%) end in Cancelled or Returned status, tying up $17,843 — 22.5% of gross revenue — in orders that were processed but never collected. Delivered orders (100) generated $19,360; Cancelled + Returned orders alone generated almost as much ($17,843) without ever completing. This is closer to a fulfillment/ops problem than a demand problem.

**3. Payment method predicts cancellation risk.**
Debit Card orders cancel or return 31.6% of the time — the highest of any payment method — versus 22.2% for PayPal, a ~9.4 point gap. Bank Transfer (25.5%) and COD (25.9%) also run well above PayPal. If this holds beyond this sample, steering customers toward PayPal at checkout (or adding fraud/validation friction to Debit Card orders) is a concrete lever to reduce the 24.8% failure rate above.

**4. "Unknown" is the largest region and among the largest payment methods in the data — a data quality problem hiding inside the dashboard.**
122 orders (24.4% of volume, $20,389 / 25.7% of revenue) have no shipping region on file — more revenue than the North region, the actual largest known region. 128 orders (25.6%) similarly have no payment method recorded. Any regional or payment strategy built on this data today is being built on a quarter of it being invisible. This is the single highest-leverage fix in the dataset: closing that intake gap would make every other regional and payment breakdown meaningfully more trustworthy.

**5. Revenue grew 13.9% year-over-year, but order volume barely moved.**
2023: $37,042 across 243 orders. 2024: $42,185 across 257 orders — a 13.9% revenue increase against only a 5.8% increase in order count. Growth is coming from higher-value orders (helped by the Website channel's $186.87 average order value, the highest of any channel), not from acquiring meaningfully more customers.

**6. Online is the largest channel by revenue, but Website has the best margin per order and the worst cancellation rate.**
Online (web + app combined under this label) drives 39.4% of revenue across 186 orders. But isolating **Website** specifically: it has the highest average order value ($186.87) of any channel and also the highest cancellation rate (19.2%, vs. 10.8% for Mobile App). High-value web orders are the most likely to fall through before completing — worth investigating what in the website checkout flow is different from the app.

## What's in the Workbook

- **Raw Sales Data** — 500 order-level records: customer, product, category, unit price, quantity, order date, ship region, payment method, order status, sales channel, discount %, and two formula-driven fields (Total Price = Unit Price × Qty; Total Sales = Total Price adjusted for Discount %)
- **ProductPriceTable** — master unit price reference used to validate the raw data via XLOOKUP
- **Six pivot sheets** (By Product, By Category, By Sales Channel, By Order Status, By Payment Method, By Region) with slicers on the Sales Channel views for interactive filtering

## Data Quality Notes

- 48 rows (9.6%) are missing an email address; 85 rows (17%) are missing a phone number — not used in any analysis above, flagged for completeness
- 66 rows (13.2%) have no discount value logged (treated as 0% here, consistent with the workbook's own Total Sales formula)
- The "Unknown" values in Ship Region and Payment Method are the largest data quality issue in the set — see Finding 4

## Recommendations

1. **Diversify beyond the Standing Desk.** With ~29% of revenue in one product, model what a 10–20% demand drop for that single SKU would do to total revenue, and prioritize promotion of the next tier of products (Monitor 24", Headphones) to reduce single-point exposure.
2. **Investigate the Cancelled/Returned pipeline before investing more in acquisition.** Recovering even a third of the $17,843 currently lost to cancellations/returns is worth more than most channel-mix optimizations in this dataset.
3. **Fix the intake process feeding Ship Region and Payment Method**, since a quarter of the data currently can't be attributed to any region or payment method — every regional/payment recommendation in this dashboard is only as reliable as this gap allows.
4. **Look at Debit Card and Website checkout flows specifically** — both show above-average cancellation/return behavior and are large enough (Debit Card: ~140+ orders touching payment; Website: 78 orders) to move the topline number if fixed.

## Tools Used

Excel — pivot tables, slicers, XLOOKUP, SUMIFS, COUNTIFS, TRIM/TEXT/DATEDIF for data cleaning and validation.


