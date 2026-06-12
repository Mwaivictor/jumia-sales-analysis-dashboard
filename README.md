 Jumia Product Performance Dashboard
### Analyzing Pricing, Discounts, and Customer Reviews

## 1. Project Overview

This is my **first-ever Excel data analysis and dashboard project**. I worked with a dataset of products scraped from **Jumia Kenya** (one of East Africa's largest e-commerce platforms) and used Excel to clean the data, enrich it with calculated fields, analyze it with pivot tables, and present the results in an interactive dashboard.

The goal of the project was to answer three business questions:

1. Do bigger discounts actually drive more customer engagement (reviews)?
2. Do higher-rated products cost more or less than poorly-rated ones?
3. Which products perform best overall when you combine price, discount, rating, and reviews?

The workbook contains three sheets:

| Sheet | Purpose |
|---|---|
| `Excel_jumia_dataset assignment` | Cleaned dataset + calculated columns + KPI summary block |
| `PIVOT` | 10 pivot tables powering the analysis (Top/Bottom 10 rankings, category counts, averages) |
| `DASHBOARD` | Interactive dashboard with 5 charts and 2 slicers |

## 2. Dataset Overview

The dataset contains **~110 home & lifestyle products** (kitchenware, storage, cleaning tools, décor, small electronics). The original scraped columns were:

| Column | Description | Raw format example |
|---|---|---|
| `Product` | Product name | `Portable Mini Cordless Car Vacuum Cleaner - Blue` |
| `Current price` | Selling price in Kenyan Shillings (KSh) | `2199` |
| `Old price` | Price before the discount | `2923` |
| `Discount` | Discount percentage | scraped as a raw percentage |
| `Review` | Number of customer reviews | scraped as `(24)` → imported as `-24` |
| `Rating` | Customer rating | text like `"4.6 out of 5"` |

A note on the data: prices range from **KSh 38 to KSh 3,750**, so this is a budget/mid-range household goods segment. Only about half the products (**58 of ~110**) have any reviews at all — a real-world data limitation I had to work around throughout the project.

## 3. Data Cleaning & Preparation

This was the most time-consuming part and taught me the most. Here is what I cleaned:

**Price formatting.** The scraped prices originally came with the "KSh" currency prefix and thousands-separator commas (e.g., `KSh 2,923`), which Excel reads as text. I stripped the currency text and commas and converted both `Current price` and `Old price` to plain numbers so they could be used in formulas.

**Negative review counts.** The review counts were scraped inside parentheses, like `(24)`. Excel interprets parentheses as negative numbers, so the whole `Review` column imported as negatives (`-24`). I fixed this by creating an `Abs_review` column using `=ABS(Review)` to recover the true count.

**Rating standardization.** Ratings came in as text strings like `"4.5 out of 5"`, which can't be averaged or charted. I extracted the numeric part into a new `Abs_rating` column (e.g., `4.5 out of 5` → `4.5`), giving ratings on a clean 0–5 scale.

**Missing values.** The dataset has genuine gaps that I documented rather than deleted:
- **52 products have no reviews or ratings** — these are simply products no customer has reviewed yet on Jumia. I kept them for the pricing/discount analysis but excluded them from rating and review calculations (Excel's `AVERAGE` ignores blanks automatically).
- **3 rows are missing the product name** but still have valid price and discount data.

**Duplicates.** After cleaning, I checked the data block and found **no duplicate product rows** — each of the ~110 products appears once.

## 4. Feature Engineering

I added several calculated columns to enrich the dataset:

**Discount amount (`Abs_dicount`)** — the absolute saving in shillings:
```
= Old price − Current price
```

**Discount percentage (`Abs_Disc(%)` and a rounded `Discount` column)**:
```
= (Old price − Current price) / Old price
```

**Rating categories (`category Rating`)** — built with a nested IF on the cleaned rating:

| Category | Rule | Count (rated products) |
|---|---|---|
| POOR | rating < 3 | 12 |
| AVERAGE | 3 ≤ rating < 4.5 | 22 |
| EXCELLENT | rating ≥ 4.5 | 23 |

*Known limitation:* my formula assigns "EXCELLENT" to the 52 unrated products by default (the IF falls through when the rating cell is blank). In the insights below I only use the **rated** products for rating-based conclusions. This is a bug I'd fix in version 2 by adding a "NOT RATED" branch first.

**Discount categories (`Discount_category`)** — nested IF on the discount percentage:

| Category | Rule | Count |
|---|---|---|
| LOW_DISCOUNT | < 20% | 18 |
| MEDIUM_DISCOUNT | 20–39% | 31 |
| HIGH_DISCOUNT | ≥ 40% | 60 |

**Performance Score** — a custom weighted score I built to rank products on more than one dimension at a time:
```
= 0.40 × Reviews + 0.35 × Rating + 0.25 × Discount %
```
Reviews get the biggest weight because they are the strongest signal of real customer engagement.

## 5. Data Analysis (Assignment Solutions)

### A. Descriptive Statistics

Computed with `AVERAGE`, `MAX`, `MIN`, `INDEX/MATCH` and `COUNT` in the KPI block at the bottom of the data sheet:

| Metric | Value |
|---|---|
| Average current price | **≈ KSh 1,187** |
| Average old price | **≈ KSh 1,811** |
| Average discount amount | **≈ KSh 624** |
| Average discount percentage | **≈ 36.7%** |
| Average rating (rated products only) | **≈ 3.89 / 5** |
| Average reviews per reviewed product | **≈ 12.7** |
| Most expensive product | **32PCS Portable Cordless Drill Set** — KSh 3,750 |
| Cheapest product | **3PCS Single Head Knitting Crochet Sweater Needle Set** — KSh 38 |
| Total products | **110** |
| Products with reviews | **58** |

### B. Trend Analysis

**Discount % vs Reviews** (scatter chart + pivot averages):

| Discount tier | Avg reviews | Avg rating |
|---|---|---|
| LOW (<20%) | 9.5 | 3.72 |
| MEDIUM (20–40%) | **15.9** | **4.25** |
| HIGH (>40%) | 10.8 | 3.66 |

The correlation between discount % and review count is **slightly negative (r ≈ −0.13)**. So **higher discounts do NOT lead to more engagement** — medium-discount products actually get the most reviews AND the best ratings, while heavily discounted products get fewer reviews and worse ratings.

**Rating vs Reviews:** the correlation is **essentially zero (r ≈ +0.06)**. Highly rated products do not automatically get more reviews. In fact, the single most-reviewed product in the dataset (69 reviews) has a poor 2.8 rating — unhappy customers leave reviews too. Review count measures *attention*, not *satisfaction*.

### C. Product Performance

**Top 10 products by discount %** (from the pivot table):

| Product | Discount |
|---|---|
| 6 In 1 Bottle Can Opener | 64% |
| Simple Metal Dog Art Sculpture | 55% |
| 5-PCS Stainless Steel Cooking Pot Set | 55% |
| LASA Folding Table Serving Stand | 55% |
| LASA 3 Tier Bamboo Shoe Bench | 54% |
| Mythco 120COB Solar Wall Light | 54% |
| 3PCS Knitting Crochet Needle Set | 53% |
| Classic Black Cat Pillow Case | 53% |
| Intelligent LED Body Sensor Night Light | 52% |
| Exfoliating Face Towel | 52% |

**Top 10 products by reviews:**

| Product | Reviews | Rating |
|---|---|---|
| 120W Cordless Handheld Vacuum Cleaner | 69 | 2.8 |
| 137 Pieces Cake Decorating Tool Set | 55 | 4.6 |
| Electronic Digital Vernier Caliper | 49 | 4.6 |
| 3D Waterproof EVA Shower Curtain | 44 | 4.6 |
| 100 Pcs Crochet Hook Tool Set | 39 | 4.7 |
| Punch-free Bathroom Storage Rack | 36 | 4.3 |
| 53 Pcs Yarn Crochet Hooks (Pansies) | 32 | 4.5 |
| Portable Mini Cordless Car Vacuum | 24 | 4.6 |
| 52 Pieces Cake Decorating Set | 20 | 4.1 |
| 53 Pcs Yarn Crochet Hooks (Fortune Cat) | 20 | 4.7 |

**Top 5 highest-rated products** (all 5.0/5, each with only 1–3 reviews): Anti-Skid Absorbent Coaster, Peacock Throw Pillow Case, LASA Aluminum Folding Hand Cart, DIY File Folder Desktop Organizer, Classic Black Cat Pillow Case.

**Top 5 lowest-rated products:** Wall-mounted Plug Fixer (2.0), 5-PCS Stainless Steel Cooking Pot Set (2.1), Electric LED UV Mosquito Killer Lamp (2.1), Artificial Potted Flowers (2.2), 7-piece Travel Storage Bag Set (2.2).

## 6. Dashboard Design

**Overview / KPI section** (data sheet): a summary block showing total products (110), average rating (3.89), average discount (36.7% / KSh 624), average price (KSh 1,187), total reviews counted, and the most/least expensive products.

**Pivot tables (10 total, on the PIVOT sheet):**
- Top 10 and Bottom 10 products by **rating**
- Top 10 and Bottom 10 products by **discount %**
- Top 10 and Bottom 10 products by **reviews**
- Product **count by discount category** (High 60 / Medium 31 / Low 18)
- Product **count by rating category** (Excellent / Average / Poor)
- **Averages by discount category** (avg reviews, avg rating, avg discount % per tier) — this pivot drives the main trend insight
- **Top 10 product performance** combining discount %, rating, reviews, and the Performance Score

**Charts (on the DASHBOARD sheet, plus working copies on the other sheets):**
- **Discount % vs Reviews Correlation** — scatter chart (trend analysis)
- **Rating vs Reviews Correlation** — scatter chart (trend analysis)
- **Top 10 Product Performance** — bar chart from the Performance Score pivot
- **Product Count by Discount %** — pie chart showing the High/Medium/Low split
- **Ranking by Product Rating** — bar chart of the best- and worst-rated products
- A 3D bar chart and additional pie/bar charts on the PIVOT sheet supporting the category breakdowns

**Slicers (interactive filters):** two slicers connected to the pivot tables —
- `Discount_category` (Low / Medium / High)
- `category Rating` (Poor / Average / Excellent)

Clicking a slicer button filters the connected pivot tables and their charts at the same time, which makes the dashboard interactive instead of static.

## 7. Key Insights

**1. Higher discounts do NOT increase customer engagement.** The discount-to-reviews correlation is slightly negative (r ≈ −0.13). The sweet spot is the **medium tier (20–40% off)**: those products average 15.9 reviews and a 4.25 rating, beating both the low tier (9.5 reviews) and the high tier (10.8 reviews, 3.66 rating).

**2. Heavy discounting often signals a weak product.** Of the rated products, **10 combine a high discount (>40%) with a poor rating (<3)** — e.g., the 5-PCS Cooking Pot Set (55% off, 2.1★) and the 120W Vacuum Cleaner (49% off, 2.8★). Big markdowns on Jumia frequently look like attempts to clear products customers don't like. Meanwhile only **one** product pairs a low discount with an excellent rating (the LED Digital Alarm Clock, 19% off, 4.6★) — good products don't need to be slashed.

**3. Ratings don't drive review volume.** The rating-to-reviews correlation is ≈ 0. The most-reviewed product (69 reviews) is also one of the worst-rated (2.8★). Reviews measure how many people *bought and reacted* — not whether they were happy.

**4. Poorly-rated products are the cheap ones.** Among rated products, POOR items average **KSh 998**, vs ~**KSh 1,370–1,400** for AVERAGE and EXCELLENT items. The very cheapest products in this segment tend to disappoint, while quality holds steady once you pass roughly KSh 1,300.

**5. The "old price" looks inflated as a marketing tactic.** The average discount is a steep 36.7%, and 55% of all products carry a "high" discount of 40%+. When more than half the catalogue is permanently 40% off, the strikethrough price is likely an anchor to make deals look bigger, rather than a real previous price.

**6. Craft & kitchen products quietly perform best.** Although the dataset has no formal category column, the product names show a pattern: crochet/knitting sets and cake-decorating kits dominate the high-engagement, high-rating zone (4 of the top 10 most-reviewed products are crochet sets rated 4.5–4.7; both cake-decorating sets are also in the top 10). Hobby and baking products appear to be a strong niche on Jumia.

**7. Data caveat:** ratings exist for only 58 of 110 products, and several "5.0★" products have just 1–3 reviews, so single ratings can dominate. These insights describe this sample, not all of Jumia.

## 8. Conclusion

As my first Excel project, this taught me far more than I expected:

- **Cleaning is most of the job.** Currency prefixes, parentheses turning numbers negative, and ratings stored as sentences all had to be fixed before a single chart could be made.
- **Categories make messy data readable.** Bucketing discounts and ratings with nested IFs turned 110 scattered rows into clear groups I could pivot on.
- **Pivot tables + slicers are powerful.** Ten pivot tables and two slicers did the heavy lifting for every ranking and breakdown in the dashboard.
- **Always question your own formulas.** My rating-category IF silently labelled blank ratings "EXCELLENT" — finding and documenting that bug was as valuable as any insight.
- **Correlation can surprise you.** I assumed bigger discounts would mean more engagement; the data showed the opposite.

Next steps for version 2: fix the blank-rating bug, add a proper product-category column, and pull a larger sample so the review-based insights are more robust.

---
*Tools used: Microsoft Excel (formulas, nested IFs, ABS, AVERAGE/MAX/MIN, pivot tables, slicers, scatter/bar/pie charts). Data: scraped from Jumia Kenya. Prices in Kenyan Shillings (KSh).*
