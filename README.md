# Task 1 – Sales Performance Analysis
RaushByte Technologies · Data Analytics Internship · Level 1

**Objective:** Analyze monthly sales and profit performance.
**Dataset:** `data_analytics_sales_dataset.xlsx` – 8,000 orders, 1 Jan 2024 to 15 May 2025, 4 regions, 5 products.
**Tools:** Python (pandas, matplotlib, openpyxl) and Excel.

## Requirements and results
| Requirement | Result |
|---|---|
| Calculate total sales | **2,017,771,508** (profit 601,497,462; profit margin 29.81%) |
| Identify top 5 selling products | 1. Headphones 425.7M · 2. Tablet 420.0M · 3. Smartwatch 410.2M · 4. Laptop 394.3M · 5. Mobile 367.5M |
| Compare region-wise sales | West 514.7M (25.5%) · East 503.9M (25.0%) · South 501.4M (24.8%) · North 497.8M (24.7%) |
| Basic charts and graphs | 4 PNG charts in `charts/`, plus 3 native Excel charts in the workbook |

(The dataset has only 5 products, so "top 5" ranks all of them.)

## Key findings
- **Monthly trend:** Sales stay in a band of roughly 104M–138M per month. The best month is **December 2024 (138.1M sales, 41.3M profit)**; the weakest full month is **February 2025 (103.8M)**. April 2024 was also low (109.0M).
- **Seasonality hint:** Sales rise from October to December 2024, dip in January–February 2025, then recover in March–April 2025. With under 17 months of data this is a pattern to watch, not a proven seasonal effect.
- **Profit margin is steady** at about 29–31% every month, so profit moves with sales rather than on its own.
- **Products:** Headphones lead, but the gap between 1st and 5th is only about 14% of the top figure. No product is dominant.
- **Regions:** West is first, but the spread between best and worst region is only 3.4% (about 17M). Regions perform almost equally.
- **Caution:** May 2025 holds only 1–15 May, so its low value is **not** a real drop.

## Data checks done
No missing values, no duplicate Order_IDs. `Sales = Quantity × Unit_Price` and `Profit = Sales − Expense` hold for all 8,000 rows. The Excel formulas were recalculated and match the Python results.

## Project structure
```
data/    original dataset
charts/  4 PNG charts
output/  Task1_Sales_Analysis.xlsx  (Data sheet + Summary sheet with SUMIFS formulas and charts)
sales_analysis.py   the whole analysis in one commented script
```

## How to run
```
pip install pandas openpyxl matplotlib
python sales_analysis.py
```

## Limitations
The dataset is synthetic-looking and uniform, so differences between regions and products are small. Findings describe this data only.
