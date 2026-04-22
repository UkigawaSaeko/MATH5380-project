# MATH 5380 — Project 1: Long-Only Multi-Asset Portfolio (Black–Litterman)

This repository contains the programming work for **Project 1** (Columbia MATH 5380): a long-only, multi-asset portfolio with Black–Litterman tilts, annual rebalancing, and a monthly backtest against a market-cap-weighted benchmark.

## Contents

| File | Description |
|------|-------------|
| `project1_multi_asset_bl.ipynb` | Main notebook: data prep, BL views, mean–variance optimization (long-only, ±5% deviation from market weights), backtest, charts, and summary statistics. |
| `build_project1_excel.py` | Python script that reads the course Excel data, reproduces the same model as the notebook, and writes `project1_results.xlsx` with **formulas** for gross/total/active return statistics (see sheet `Data_Monthly`). |
| `project1_results.xlsx` | Generated output (re-run the script after changing the model or data). |
| `Project 1.pdf` | Official assignment instructions (report, Excel, and code requirements). |

## Data

Place **`Data for final project 1.xlsx`** (provided by the instructor) in the **parent** folder of this project, *or* update the path in the notebook and in `build_project1_excel.py` (see `candidates` / `excel_path` logic).

Required sheets: **`Index returns in USD`**, **`Market values in USD`**.

## Environment

- **Python 3.10+** recommended (3.12 used in development).  
- Packages: `pandas`, `numpy`, `scipy`, `matplotlib`, `openpyxl` (for Excel export).  
- A **NumPy 1.26.x** line is often easier with Anaconda binary stacks; NumPy 2.x may require matching wheels for all compiled dependencies.

## Investment Universe

Our investment universe satisfies the diversification requirement by including both equity and fixed-income categories. Specifically, it spans U.S., developed ex-U.S., and emerging-market equities across value, growth, and small-cap styles, as well as several bond sectors including U.S. aggregate, global aggregate, inflation-linked, municipal, and high-yield bonds. For each rebalancing year, benchmark weights are derived from the provided market value dataset.

## Rationale for the Two View Portfolios

Our two views are motivated by a coherent macroeconomic and market-structure narrative.

On the equity side, we impose a mild preference for **U.S. growth equities** over **non-U.S. developed growth equities**. This view reflects the relative depth, liquidity, and information efficiency of the U.S. market, which may provide a structural advantage for U.S. growth stocks compared with their developed ex-U.S. counterparts.

On the fixed-income side, we impose a mild preference for **U.S. aggregate bonds** over **global high-yield bonds**. This reflects a more defensive allocation preference toward higher-quality duration exposure rather than credit-sensitive cyclical beta, especially in an environment where risk repricing and liquidity conditions may matter.

These two views are designed to be **complementary across asset classes**. Instead of concentrating multiple views on the same style dimension, we introduce one view in equities and one view in fixed income, which helps maintain diversification in the source of active bets.

## Covariance Estimation

For each annual rebalancing date, we estimate the covariance matrix of asset returns using the most recent 36 months of prior monthly returns. The sample covariance matrix is annualized by multiplying by 12. This estimation is repeated every year using only information available at the rebalancing date, which avoids look-ahead bias.

Because the full 17×17 covariance matrix is large and changes every year, we report a representative covariance (or correlation) heatmap for one rebalancing year in the main text, while the full yearly matrices are available upon request / in the appendix.

## How to run

1. **Notebook**  
   Open `project1_multi_asset_bl.ipynb` in Jupyter or VS Code, select your conda/virtual environment, then run all cells from the top (after pointing to the data file if needed).

2. **Excel workbook** (from the project root directory)  
   ```bash
   python build_project1_excel.py
   ```  
   This overwrites or creates `project1_results.xlsx`. Close the file in Excel if you see a permission error.

## Model summary (for orientation)

- **Universe:** intersection of return columns and market-value columns; names aligned to the returns file (see notebook name mapping).  
- **Views:** two relative-return views (see notebook section “参数与观点”); `P`, `Q`, and BL parameters `tau`, `omega_scale` are set there.  
- **Each year:** 36+ months of history up to the rebalance date → annualized covariance → implied `π` → Black–Litterman `μ_BL` → optimal weights with **Σ**, **μ_BL**, and the same risk-aversion as in `π`.  
- **Benchmark in code and in Excel:** within each **calendar year of performance**, the benchmark uses the **year-end market weights** chosen at the prior year’s rebalance (constant for that year’s months). The course PDF also describes an alternative “prior-month cap weights × current returns”; if you use that, document it consistently in the report.  
- **No look-ahead:** estimation window ends on or before the rebalance date; the performance year is strictly after that date (assertions in the notebook loop).

## Submission (per `Project 1.pdf`)

The course requires a **short report (≤5 pages)**, a **spreadsheet** with key results and **formulas**, and **code**. This repo supplies the code and a template Excel builder; the written report is separate.

## License / use

Course project use only; data files are property of the instructor.
