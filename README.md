# Sales Performance Analysis

A beginner-friendly data analyst portfolio project that transforms 1,478 retail transactions into a complete analysis with an Excel dashboard and interactive web app.

## What's in this project?

- **Excel Dashboard** (`Sales_Analysis_Dashboard.xlsx`) — Professional 4-sheet workbook with formulas, tables, charts, and conditional formatting
- **Streamlit App** (`dashboard.py`) — Interactive web dashboard with filters for date, location, product, and customer type
- **Python Scripts** — Clean analysis and dashboard generation code

## Quick stats

| Metric | Value |
|--------|-------|
| Transactions | 1,478 |
| Units sold | 4,409 |
| Gross sales | 978,834.19 |
| Discounts | 13,483.56 |
| Net revenue | 965,350.63 |
| Date range | Jan 2023 – Dec 2024 |

## Top performers

- **By units sold:** Wireless Charger (300 units)
- **By revenue:** Laptop (317,226.76)
- **Best location:** Berlin (216,117.45)

## How to run

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run analysis and generate Excel dashboard:
```bash
python analyze_sales.py
python create_excel_dashboard.py
```

3. Launch Streamlit app:
```bash
streamlit run dashboard.py
```

## File structure

```
Project 1/
├── Sales_Analysis_Dashboard.xlsx   # Excel workbook (4 sheets)
├── create_excel_dashboard.py      # Excel dashboard generator
├── analyze_sales.py               # Data analysis script
├── dashboard.py                   # Streamlit web app
├── requirements.txt               # Python dependencies
└── output/                        # Generated files
```

## Source

Dataset from Sven Bosau ([pythonandvba.com](https://pythonandvba.com)).
