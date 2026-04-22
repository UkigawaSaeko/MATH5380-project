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
