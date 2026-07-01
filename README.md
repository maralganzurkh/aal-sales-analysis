# AAL Q4 2020 Sales Analysis

**Python data analysis project | SimpliLearn Microsoft AI Engineer Program**

A complete end-to-end sales analysis of Australian Apparel Limited (AAL) Q4 2020 retail data, built using Python and pandas.

---

## What This Project Does

- Loads and cleans AAL Q4 2020 sales CSV data
- Computes descriptive statistics (mean, median, mode, std dev) for sales and units
- Aggregates sales by state and customer group
- Builds daily, weekly, monthly, and quarterly sales time series
- Analyses sales by time-of-day categories
- Exports a multi-sheet Excel report with all key tables
- Generates bar charts, line charts, and box plots using matplotlib

## Key Findings

- Identifies top and bottom performing states by total sales
- Reveals strongest and weakest customer groups per underperforming state
- Visualises sales trends across the quarter

## Tech Stack

| Tool | Purpose |
|---|---|
| Python 3 | Core language |
| pandas | Data loading, cleaning, aggregation |
| numpy | Normalisation and statistical calculations |
| matplotlib | Charts and visualisations |
| xlsxwriter | Excel report generation |

## How to Run

```bash
# 1. Clone the repo
git clone https://github.com/maralganzurkh/aal-sales-analysis

# 2. Install dependencies
pip install pandas numpy matplotlib xlsxwriter

# 3. Place AusApparalSales4thQrt2020.csv in the same folder

# 4. Run
python aal_q4_sales_analysis.py
```

## Output

- `AAL_Q4_2020_Sales_Report.xlsx` — multi-sheet Excel report
- Charts displayed inline during execution

---

*Project completed as part of the Microsoft AI Engineer Program — SimpliLearn (2026)*  
*Author: Maral-Od Ganzurkh | [maralganzurkh.com](https://maralganzurkh.com)*
