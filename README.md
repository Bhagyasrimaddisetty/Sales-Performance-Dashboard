# 📊 Sales Performance Dashboard

An interactive Excel dashboard analyzing **1,200+ sales transactions** across regions, product categories, customers, and monthly trends for FY 2024.

> Built as a portfolio project to demonstrate data analysis, business intelligence, and dashboard design skills.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| Microsoft Excel | KPI cards, charts, conditional formatting, pivot tables |
| Power BI | Connect to `data/sales_data.csv` for live interactive visuals |
| Python (pandas, openpyxl) | Data generation & Excel workbook automation |

---

## ✨ Features

- 💰 **Revenue & Profit KPI Cards** — Total Revenue, Profit, Orders, Active Regions
- 📅 **Monthly Sales Trends** — Line chart with MoM growth % (Jan–Dec 2024)
- 🌍 **Region-wise Sales** — Bar chart + Region × Category pivot table
- 🛍️ **Product Category Performance** — Pie chart + Top 10 products ranking
- 👑 **Top Customers** — 🥇🥈🥉 medal rankings with revenue share %
- 💹 **Profit Analysis** — Channel profitability + gross margin breakdown

---

## 📂 Project Structure

```
sales-performance-dashboard/
│
├── Sales_Performance_Dashboard.xlsx   # Main Excel dashboard (7 sheets)
├── data/
│   └── sales_data.csv                 # Raw dataset — 1,267 transactions
├── scripts/
│   └── generate_data.py               # Python script to regenerate workbook
└── README.md
```

---

## 📊 Dashboard Sheets

| Sheet | Description |
|---|---|
| 📊 Dashboard | Executive summary — KPI cards + 3 embedded charts |
| 📋 Raw Data | Full transactional dataset (1,267 records) |
| 📅 Monthly Trends | MoM revenue/profit with color-scale formatting |
| 🌍 Region Analysis | Data bars + Region × Category pivot |
| 🛍️ Product Categories | Category KPIs + Top 10 products |
| 👑 Top Customers | Medal rankings + revenue share |
| 💹 Profit Analysis | Channel profitability + margin summary |

---

## 📈 Key Metrics (FY 2024)

- **Revenue:** $3.2M+ across 5 regions
- **Profit Margin:** ~38–42%
- **Orders:** 1,267 transactions
- **Categories:** Electronics, Clothing, Home & Garden, Sports, Books
- **Regions:** North, South, East, West, Central
- **Channels:** Online, Retail Store, Wholesale, Direct Sales

---

## ⚡ Quick Start

### Option 1 — Excel (No Setup Required)
1. Download `Sales_Performance_Dashboard.xlsx`
2. Open in Microsoft Excel 2016+
3. Navigate sheets using tabs at the bottom

### Option 2 — Power BI
1. Open Power BI Desktop
2. **Get Data → Text/CSV** → select `data/sales_data.csv`
3. Build visuals using the pre-cleaned dataset

### Option 3 — Regenerate with Python
```bash
pip install pandas openpyxl
python scripts/generate_data.py
```

---

## 🔢 Power BI DAX Measures

```dax
Total Revenue = SUM(sales_data[Revenue])

Total Profit = SUM(sales_data[Profit])

Profit Margin % = DIVIDE(SUM(sales_data[Profit]), SUM(sales_data[Revenue]))

MoM Revenue Growth =
    VAR CurrentRevenue = [Total Revenue]
    VAR PreviousRevenue =
        CALCULATE([Total Revenue], DATEADD('sales_data'[Date], -1, MONTH))
    RETURN DIVIDE(CurrentRevenue - PreviousRevenue, PreviousRevenue)
```

---

## 👩‍💻 Author

**Bhagya Sri Maddisetty**
B.Tech CSE (AI/ML) | Mohan Babu University | CGPA: 9.3/10

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://www.linkedin.com/in/bhagya-sri-maddisetty-064102305/)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-black?logo=github)](https://github.com/Bhagyasrimaddisetty)

---

## 📄 License

MIT License — free to use and modify for portfolio purposes.
