# 📊 Analyzer-Excel

### Excel Data Analysis & Dashboard Project


## 🔍 Overview

**PR. 2 Analyzer** is a hands-on Excel project that walks through a complete data analysis workflow on a customer sales dataset. It covers everything from highlighting top customers and simulating discount scenarios to running statistical regression and summarizing results in a clean, interactive dashboard.

The goal is to turn raw transactional data into **clear, decision-ready insights** using only Microsoft Excel's built-in tools.

---

## 🎯 Objectives

- Identify and visualize **top-performing customers**
- Measure the **impact of discounts on profit** with scenario modelling
- Understand the **relationship between Sales and Profit** using regression
- Summarize the dataset with **descriptive statistics**
- Track **monthly sales growth** with visual indicators
- Build a **pivot-based dashboard** for quick business reporting

---

## 🧰 Topics Covered

| # | Topic | Category |
|---|-------|----------|
| 1 | Conditional Formatting | Visualization |
| 2 | What-If Analysis | Scenario Modelling |
| 3 | Linear Regression (Data Analysis ToolPak) | Statistics |
| 4 | Data Analysis Feature (Descriptive Stats, Regression, Histogram) | Statistics |
| 5 | Data Storytelling | Communication |
| 6 | Up/Down Arrows & Symbols in Custom Formatting | Formatting |
| 7 | Timestamp Creation | Functions |
| 8 | High-Value Customers using `GROUPBY()`, `INDEX()`, `MATCH()` | Lookup |
| 9 | Advanced Descriptive Analytics | Analytics |
| 10 | Pivot Table with Dashboard | Reporting |
| 11 | Bar, Line & Pie Charts | Visualization |

---

## ✅ Tasks & Solutions

| # | Task | Approach |
|---|------|----------|
| 1 | Highlight **Top 10 customers** by total purchase | Conditional Formatting → *Top/Bottom Rules* → Top 10 Items |
| 2 | **What-If Analysis** – impact of discount on total profit | Data Table / Scenario Manager / Goal Seek on the discount input |
| 3 | **Linear Regression** (Profit vs Sales) | Data → Data Analysis → Regression |
| 4 | **Descriptive Statistics** on the dataset | Data → Data Analysis → Descriptive Statistics |
| 5 | **Up/Down arrows** on monthly sales growth | Custom number format with ▲ ▼ symbols + color rules |
| 6 | **Timestamp** column | `=NOW()` |
| 7 | Identify **high-value customers** | `INDEX()` + `MATCH()` + Filters |
| 8 | **Pivot Table** – total sales by region & product | Insert → PivotTable |
| 9 | **Bar, Line & Pie charts** for KPIs | Insert → Charts |
| 10 | **Dashboard** summarizing insights | Pivot charts + slicers + KPI cards |

---

## 🧮 Key Formulas Used

```excel
' Timestamp
=NOW()

' Monthly growth %
=(Current_Month - Previous_Month) / Previous_Month

' Find the top customer's name
=INDEX(Customer_Range, MATCH(MAX(Total_Purchase_Range), Total_Purchase_Range, 0))

' Total profit after discount (What-If model)
=Sales * (1 - Discount_Rate) - Cost

' Aggregate total purchase per customer
=GROUPBY(Customer_Range, Purchase_Range, SUM)
```

### 🔣 Custom Format for Growth Arrows

```text
[Color10]▲ 0.0%;[Red]▼ 0.0%;[Blue]● 0.0%
```

> 💡 Green ▲ for growth, Red ▼ for decline, Blue ● for no change.

---

## 💡 Key Insights

> Replace the placeholders below with your actual findings.

- 🏆 **Top 10 customers** contribute **X%** of total revenue.
- 📉 Increasing the discount from **A% to B%** changes total profit by **₹ / $ X**.
- 📈 Regression shows **Sales** explains **R² = X** of the variation in **Profit**.
- 🌍 **[Region]** and **[Product]** deliver the highest sales.
- 📅 Sales peaked in **[Month]** and dipped in **[Month]**.

---

## 📂 Project Structure

```text
PR2-Analyzer/
│
├── 📄 PR2_Analyzer.xlsx      # Main workbook (all sheets)
├── 📄 README.md              # Project documentation
└── 📁 images/                # Screenshots used in README
    ├── dashboard.png
    ├── pivot-table.png
    ├── regression.png
    └── charts.png
```

### 📑 Workbook Sheets

| Sheet | Purpose |
|-------|---------|
| `Data` | Raw dataset |
| `Conditional Formatting` | Top 10 customers highlighted |
| `What-If` | Discount vs. profit scenarios |
| `Regression` | ToolPak regression output |
| `Descriptive Stats` | Summary statistics & histogram |
| `Growth` | Monthly growth with ▲ ▼ indicators |
| `High-Value Customers` | INDEX/MATCH lookup & filters |
| `Pivot` | Sales by region and product |
| `Dashboard` | Final KPI dashboard |

---

## 🎬 Project Demonstration
  <a href="">
    <img src="https://img.shields.io/badge/🎬%20Project%20Video-success?style=for-the-badge">
  </a>
</p>

---

