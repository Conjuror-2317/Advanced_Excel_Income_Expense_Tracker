# 📊 Personal Income & Expense Tracker — Excel

A multi-sheet personal finance tracker built in Excel to analyse 12 months of income and expenses. Demonstrates advanced Excel skills including SUMIFS, INDEX-MATCH, nested IF logic, VLOOKUP cross-sheet referencing, pivot tables, and dynamic dashboards.

---

## 📸 Preview

> **How to add your screenshots:**  
> Take a screenshot of each sheet → save to an `images/` folder in this repo → replace the placeholder paths below.

### Summary sheet
![Summary sheet showing monthly income vs expenses and category breakdown](images/summary.png)

### Charts
![Monthly income vs expenses bar chart and category spend pie chart](images/charts.png)

### Budget vs actual
![Budget tracker comparing planned spend to actual spend by category](images/budgets.png)

### Pivot tables
![Pivot table showing monthly breakdown with slicer filters](images/pivot1.png)
![Pivot table showing category-wise expense breakdown with filter](images/pivot2.png)

---

## 🗂️ File structure

| Sheet | Description |
|---|---|
| `DATA` | 42 raw transactions across 12 months — Date, Description, Category, Type, Amount |
| `SUMMARY` | Formula-driven monthly summary with status flags, category breakdown, and key stats |
| `CHART` | Monthly income vs expenses bar chart and category spend pie chart |
| `BUDGETS` | Annual budget vs actual spend by category using cross-sheet VLOOKUP |
| `PIVOT1` | Monthly income vs expenses pivot table (same output as SUMIFS — built both to compare approaches) |
| `PIVOT2` | Category spend breakdown pivot table with slicer for dynamic filtering |

---

## 📐 Excel skills demonstrated

### Lookup & reference
- `INDEX-MATCH` — identifies highest and lowest savings month by value lookup (handles ties using `TEXTJOIN + IF` array formula)
- `VLOOKUP` — pulls budget figures from the `BUDGETS` sheet into the summary cross-sheet reference (`Budgets!$A$2:$B$7`)

### Conditional aggregation
- `SUMIFS` — aggregates income and expenses by month and type simultaneously using `DATE()` boundary conditions
- `COUNTIFS` / `AVERAGEIFS` — transaction count and average spend per category

### Logic
- Nested `IF` — 4-level savings health flag per month: `Deficit → Low savings → Moderate → Healthy`, calculated against savings rate thresholds

### Error handling
- `IFERROR` — wraps VLOOKUP calls to return `"Not found"` cleanly for missing categories

### Data tools
- Pivot tables — two pivot tables replicating SUMIFS outputs dynamically; demonstrates both formula and pivot approaches to the same problem
- Slicers — dynamic filtering by Type and Category applied to pivot tables
- Data validation — dropdown menus on `Category` and `Type` columns to enforce consistent data entry
- Power Query — used to clean and load source data

### What-if analysis
- Goal Seek — used to calculate income required to hit a 30% savings rate
- 2-variable Data Table — models net savings across combinations of income growth and expense ratio

### Formatting
- Conditional formatting with colour scales — green/white/red applied to Net Savings column
- Conditional formatting with rules — red flag when savings rate < 0%, budget overspend highlighted red
- Number formatting — comma separators, 2 decimal places, percentage display

---

## 📊 Key insights from the data

| Metric | Value |
|---|---|
| Total annual income | ₹4,61,500 |
| Total annual expenses | ₹70,331.75 |
| Net annual savings | ₹3,91,168.25 |
| Savings rate | ~84.8% |
| Highest savings month | May (₹53,930) |
| Lowest savings month | November (–₹6,500, deficit) |
| Deficit months | September, November |
| Largest expense category | Food (₹19,463) |

---

## 🛠️ Tools used

- Microsoft Excel 2021
- Power Query (data cleaning)

---

## 💡 Notes

- Data used in this project is sample/dummy data created for practice purposes
- The same monthly breakdown is built two ways — using `SUMIFS` formulas and using pivot tables — to demonstrate both approaches and understand their trade-offs
- Pivot tables update dynamically when new data is added to the `DATA` sheet (right-click → Refresh)

---

*Part of my [finance portfolio](https://github.com/your-username/finance-portfolio) — built as Week 1 of a 17-week financial analyst roadmap.*
