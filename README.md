# 📊 Superstore Sales Excel Dashboard

> **An end-to-end interactive Excel dashboard built on a Superstore Sales 2024 dataset — covering data import, cleaning, pivot analysis, advanced formulas, visualizations, and a fully dynamic KPI dashboard.**

---

## 🗂️ Project Overview

| 🏷️ Attribute | 📋 Details |
|---|---|
| 📁 **File Name** | `RENSEE_GAJIPARA_ExcelDashboard.xlsx` |
| 🗃️ **Dataset** | Superstore Sales 2024 (Kaggle-style, 60 records) |
| 📅 **Year** | 2024 |
| ⏱️ **Time Limit** | 3 Hours |
| 👤 **Author** | **RENSEE GAJIPARA** |

---

## 📦 Dataset — Superstore Sales 2024

The dataset is **embedded directly** in the `Raw_Data` sheet — no external file required.

### 📋 Dataset Columns

| Column | Description | Type |
|---|---|---|
| `Order_ID` | Unique order identifier (ORD-1001 … ORD-1060) | Text |
| `Order_Date` | Date order was placed (YYYY-MM-DD) | Date |
| `Ship_Date` | Date order was shipped | Date |
| `Region` | East / West / North / South | Text |
| `Category` | Technology / Furniture / Office Supplies | Text |
| `Sub_Category` | Laptops, Phones, Chairs, Paper… | Text |
| `Product_Name` | Specific product name | Text |
| `Sales` | Revenue generated ($) | Number |
| `Quantity` | Units sold | Integer |
| `Discount` | Discount rate applied (0 – 0.30) | Percentage |
| `Profit` | Net profit/loss ($) | Number |

> 📌 Data spans **4 regions**, **3 categories**, **12 sub-categories**, **60 unique orders**

---

## 📁 Workbook Structure — 7 Sheets

```
RENSEE_GAJIPARA_ExcelDashboard.xlsx
├── 📄 Raw_Data           → Source dataset (60 rows, Excel Table)
├── 🧹 Data_Cleaning      → Audit report & data quality metrics
├── 📊 Pivot_Tables       → Region / Category / Monthly summaries
├── 🔢 Advanced_Formulas  → 21 formula demos with live results
├── 📈 Data_Visualization → 3 charts + conditional formatting
├── 🖥️  Dashboard          → Interactive KPI dashboard
└── 📋 Documentation      → Sheet index, instructions, naming guide
```

---

## 🔢 Formulas Used — Complete Reference

### ✅ Task 1 — Data Import & Cleaning

| Formula | Purpose |
|---|---|
| `=COUNTA(Raw_Data!A:A)-1` | Count total records |
| `=COUNTBLANK(Raw_Data!H2:H61)` | Check for missing Sales values |
| `=SUMPRODUCT((COUNTIF(A2:A61,A2:A61)-1))` | Detect duplicate Order IDs |
| `=MIN() / MAX() / AVERAGE()` | Basic stats on Sales |
| `=SUMPRODUCT(1/COUNTIF(D2:D61,D2:D61))` | Count unique regions |

---

### ✅ Task 2 — Pivot Table Formulas

| Formula | Purpose |
|---|---|
| `=SUMIF(region_col, "East", sales_col)` | Total Sales for a region |
| `=AVERAGEIF(region_col, "East", disc_col)` | Avg discount per region |
| `=COUNTIF(cat_col, "Technology")` | Count orders per category |
| `=SUMPRODUCT((MONTH(DATEVALUE(date_col))=6)*sales_col)` | Monthly sales (June) |
| Running totals: `=I{r-1}+H{r}` | Cumulative monthly sales |

---

### ✅ Task 3 — Data Visualization Formulas

| Formula / Feature | Used For |
|---|---|
| `=SUMIF(region_col, "East", sales_col)` | Bar chart data source |
| `=SUMIF(cat_col, "Technology", sales_col)` | Pie chart data source |
| `=SUMPRODUCT(MONTH(...)=mn * sales)` | Line chart monthly data |
| **Color Scale** (Red→Yellow→Green) | Sales conditional formatting |
| **Data Bars** (Blue) | Profit visual bars |
| **Color Scale** (White→Orange) | Discount gradient |

---

### ✅ Task 4 — Advanced Formulas (21 Demos)

| # | Formula Type | Example | Purpose |
|---|---|---|---|
| 1 | **VLOOKUP** | `=VLOOKUP("ORD-1001",Raw_Data!A:K,8,FALSE)` | Lookup sales by Order ID |
| 2 | **INDEX-MATCH** | `=INDEX(G:G,MATCH(MAX(H:H),H:H,0))` | Top-selling product name |
| 3 | **IF** | `=IF(AVERAGE(K2:K61)>0,"Profitable","Loss")` | Profitability status |
| 4 | **Nested IF** | `=IF(MAX>3000,"High",IF(MAX>1500,"Mid","Low"))` | Sales volume category |
| 5 | **SUMIFS** | `=SUMIFS(H:H,D:D,"East",E:E,"Technology")` | Multi-condition sum |
| 6 | **COUNTIFS** | `=COUNTIFS(D:D,"West",K:K,">0")` | Multi-condition count |
| 7 | **AVERAGEIFS** | `=AVERAGEIFS(H:H,E:E,"Furniture",J:J,">0")` | Conditional average |
| 8 | **IFERROR** | `=IFERROR(VLOOKUP(...),"Not Found")` | Error-safe lookup |
| 9 | **TEXT** | `=TEXT(DATEVALUE("2024-06-15"),"MMMM YYYY")` | Date formatting |
| 10 | **LEN & TRIM** | `=LEN(TRIM(G2))` | Clean character count |
| 11 | **Concatenate** | `=D2&" \| "&E2&" \| "&F2` | Build label string |
| 12 | **RANK** | `=RANK(H2,$H$2:$H$61,0)` | Rank by sales |
| 13 | **LARGE** | `=LARGE(H2:H61,1)` | Highest sale value |
| 14 | **SMALL** | `=SMALL(H2:H61,1)` | Lowest sale value |
| 15 | **PERCENTILE** | `=PERCENTILE(H2:H61,0.75)` | 75th percentile |
| 16 | **STDEV** | `=STDEV(H2:H61)` | Sales std deviation |
| 17 | **CORREL** | `=CORREL(H2:H61,K2:K61)` | Sales vs Profit correlation |
| 18 | **SUMPRODUCT** | `=SUMPRODUCT(H2:H61*I2:I61)` | Weighted sum |
| 19 | **Unique Count** | `=SUMPRODUCT(1/COUNTIF(D2:D61,D2:D61))` | Distinct region count |
| 20 | **MAX (conditional)** | `=SUMPRODUCT(MAX((E2:E61="Technology")*H2:H61))` | Max sales by category |
| 21 | **MIN (conditional)** | `=MIN(IF(K2:K61>0,K2:K61))` | Min positive profit |

---

### ✅ Task 5 — Dashboard Formulas

| Formula | Dashboard Card |
|---|---|
| `=SUM(Raw_Data!H2:H61)` | 💰 Total Sales KPI |
| `=SUM(Raw_Data!K2:K61)` | 📈 Total Profit KPI |
| `=COUNTA(Raw_Data!A2:A61)` | 🛒 Total Orders KPI |
| `=SUM(Raw_Data!I2:I61)` | 📦 Total Quantity KPI |
| `=AVERAGE(Raw_Data!H2:H61)` | 💹 Avg Sale KPI |
| `=B{r}/SUM($B$9:$B$12)` | Region % Share |
| `=IF(H{r}=0,0,I{r}/H{r})` | Category Profit % |
| `=INDEX(G:G,MATCH(LARGE(H:H,1),H:H,0))` | Top-1 Product Name |
| `=SUM(K2:K61)/SUM(H2:H61)` | Overall Profit Margin % |
| `=COUNTIF(J2:J61,">0")` | Discounted Orders count |
| `=AVERAGEIF(J2:J61,">0",J2:J61)` | Avg discount on disc. orders |

---

## 🖥️ Dashboard Preview

> **The Dashboard sheet** provides a complete at-a-glance view of the 2024 Superstore Sales performance:

```
┌─────────────────────────────────────────────────────────────────────┐
│  📊 SUPERSTORE SALES PERFORMANCE DASHBOARD  |  2024                 │
├──────────────┬──────────────┬──────────────┬──────────────┬─────────┤
│ 💰 Total     │ 📈 Total     │ 🛒 Total     │ 📦 Total    │ 💹 Avg  │
│ Sales        │ Profit       │ Orders       │ Qty         │ Sale    │
│ [Dynamic]    │ [Dynamic]    │ [Dynamic]    │ [Dynamic]   │[Dynamic]│
├──────────────┴──────────────┴──────────────┴──────────────┴─────────┤
│  📌 Sales by Region    │   🏷️ Sales by Category                     │
│  East / West / N / S   │   Technology / Furniture / Office Supplies │
├────────────────────────┴────────────────────────────────────────────┤
│  🏆 Top 5 Products by Sales    │  💡 Profit Analysis               │
│  INDEX-MATCH + LARGE formulas  │  Max, Min, Avg, Margin %          │
├──────────────────────────────── ────────────────────────────────────┤
│  🔖 Discount Impact Analysis                                        │
│  Discounted vs Non-Discounted sales, Avg discount rate              │
└─────────────────────────────────────────────────────────────────────┘
```

> 📌 **All values are formula-driven** — editing `Raw_Data` automatically updates every KPI, table, and analysis section.

---

## 🚀 How to Use

```
1. Open RENSEE_GAJIPARA_ExcelDashboard.xlsx in Excel (2016 or later)
2. Navigate to → Dashboard (start here for overview)
3. Explore → Pivot_Tables for regional/category/monthly breakdowns
4. Study → Advanced_Formulas for live formula demos
5. Review → Data_Visualization for charts & conditional formatting
6. Check → Data_Cleaning for data quality metrics
7. Reference → Documentation for full sheet index & instructions
```

---

## 📋 Project Requirements Checklist

| ✅ | Task | Sheet |
|---|---|---|
| ✅ | Dataset Selection & Import | `Raw_Data` |
| ✅ | Data Cleaning & Preparation | `Data_Cleaning` |
| ✅ | Create Pivot Tables | `Pivot_Tables` |
| ✅ | Advanced Formulas (VLOOKUP, INDEX-MATCH, IF, nested) | `Advanced_Formulas` |
| ✅ | Data Visualization (Charts + Conditional Formatting) | `Data_Visualization` |
| ✅ | Interactive Dashboard (Slicers-style, Dynamic ranges) | `Dashboard` |
| ✅ | Documentation (Naming conventions, Instructions) | `Documentation` |

---

## 🛠️ Technical Details

- **Tool:** Microsoft Excel 2016+ / Excel 365
- **Language:** Excel Formulas (no VBA/macros)
- **Total Formulas:** 272 live formulas across all sheets
- **Dataset Rows:** 60 orders, 11 columns
- **Charts:** Bar (Sales by Region), Pie (Sales by Category), Line (Monthly Trend)
- **Conditional Formatting:** Color Scale, Data Bars, Color Gradient
- **File Name Convention:** `RENSEE_GAJIPARA_ExcelDashboard.xlsx`

---

## 👤 Author

<div align="center">

### 🎓 RENSEE GAJIPARA

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://www.microsoft.com/excel)
[![Data Analysis](https://img.shields.io/badge/Data_Analysis-2E75B6?style=for-the-badge&logo=databricks&logoColor=white)](#)
[![Dashboard](https://img.shields.io/badge/Dashboard-ED7D31?style=for-the-badge&logo=powerbi&logoColor=white)](#)

> 📊 **Excel Dashboard Project 2024**
> 🗃️ **Dataset:** Superstore Sales 2024
> 📁 **Submission:** `RENSEE_GAJIPARA_ExcelDashboard.xlsx`

---

*Built with ❤️ using Microsoft Excel | Data sourced from Kaggle-style Superstore dataset*

</div>
