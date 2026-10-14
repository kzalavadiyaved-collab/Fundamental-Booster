<div align="center">

# 📊 Fundamental Booster

### Excel Formulas Mastery Project — PR. 1

*Live, formula-driven workbook covering logic, lookups, dynamic references, dates and filtering.*

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Formulas](https://img.shields.io/badge/Formulas-327-blue?style=for-the-badge)
![Errors](https://img.shields.io/badge/Errors-0-brightgreen?style=for-the-badge)
![Sheets](https://img.shields.io/badge/Sheets-8-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

[Overview](#-overview) •
[Structure](#-workbook-structure) •
[Task Coverage](#-task-coverage) •
[Quick Start](#-quick-start) •
[Assumptions](#-assumptions) •
[Compatibility](#-compatibility)

</div>

---

## 🎯 Overview

This project applies core Excel functions to realistic **student**, **sales** and **employee** datasets. Every result is a live formula (no hardcoded values), so changing any input updates the whole workbook instantly.

| | |
|---|---|
| **Deliverable** | `Fundamental_Booster.xlsx` |
| **Formulas** | 327, all calculating without errors |
| **Data** | 12 students · 20 sales · 10 employees · 5 products |
| **Skill level** | Beginner → Intermediate |

### ✨ Highlights

- 🔗 Relative, absolute and mixed references used correctly
- 🧠 `IF`, nested `IF`, `IF + AND`, `IF + OR` decision logic
- 📈 Conditional aggregation with `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`
- 🔍 `VLOOKUP`, `INDEX/MATCH`, `XLOOKUP`, `XMATCH` lookups
- 🔤 Text cleaning with `LEFT`, `FIND`, `UPPER`, `LOWER`
- 🧭 Dynamic ranges with `INDIRECT` and `OFFSET`
- 📅 Date maths with `DATEDIF` and `TODAY`
- 🧮 Financial rounding with `ROUND`, `CEILING`, `FLOOR`
- 🧾 Multi-value output with `FILTER`

---

## 🗂 Workbook Structure

```text
Fundamental_Booster.xlsx
│
├── 📋 Analysis        → inputs (yellow) + one live result per task
├── 🎓 Students        → scores, average, grade, AND-check, filter helper
├── 🏷  Products        → code, name, unit price (lookup table)
├── 💰 Sales           → 20 transactions, discounts, eligibility, tax
├── 👥 Employees       → joining date, DOB, salary, text functions, age
├── 📆 Monthly         → sales per salesperson per month (SUMIFS)
├── 🏆 TopPerformers   → dynamic list of students above a threshold
└── 📐 References      → absolute / relative / mixed reference demo
```

---

## ✅ Task Coverage

### 1️⃣ References, Formatting & Logic

| Task | Sheet | Formula |
|---|---|---|
| Absolute vs relative references | References, Sales!M | `=L2*(1+$P$2)` |
| Classify student grades | Students!G | `=IF(F2>=90,"A+",IF(F2>=80,"A",IF(F2>=70,"B",...)))` |
| Discount by price threshold | Sales!I | `=IF(F2>=500,0.15,IF(F2>=100,0.1,IF(F2>=50,0.05,0)))` |
| Math **and** Science above 80 | Students!H | `=IF(AND(C2>80,D2>80),"Yes","No")` |
| Product discount eligibility | Sales!K | `=IF(OR(F2>=500,G2>=10),"Eligible","Not Eligible")` |

### 2️⃣ Conditional Aggregation

| Task | Sheet | Formula |
|---|---|---|
| Students with Math above 50 | Analysis | `=COUNTIFS(Students!C2:C13,">50")` |
| Sales for a region and product | Analysis | `=SUMIFS(Sales!H2:H21,Sales!C2:C21,B3,Sales!D2:D21,B4)` |
| Average of scores above 60 | Analysis | `=AVERAGEIFS(Students!F2:F13,Students!F2:F13,">60")` |

### 3️⃣ Lookup & Reference

| Task | Sheet | Formula |
|---|---|---|
| Student name by ID | Analysis | `=VLOOKUP(B5,Students!A2:B13,2,FALSE)` |
| Product price by code | Sales!F | `=VLOOKUP(E2,Products!$A$2:$C$6,3,FALSE)` |
| Sales by person and month | Analysis | `=INDEX(Monthly!B2:D5,MATCH(B6,Monthly!A2:A5,0),MATCH(B7,Monthly!B1:D1,0))` |
| Employee detail, any column | Analysis | `=INDEX(Employees!A2:K11,MATCH(B9,Employees!A2:A11,0),MATCH(B10,Employees!A1:K1,0))` |
| Salary without sorted data | Analysis | `=XLOOKUP(B9,Employees!A2:A11,Employees!F2:F11,"Not found")` |
| Position of a product | Analysis | `=XMATCH(B4,Sales!D2:D21)` |

### 4️⃣ Text, Dynamic & Date Functions

| Task | Sheet | Formula |
|---|---|---|
| First name from full name | Employees!G | `=LEFT(B2,FIND(" ",B2)-1)` |
| Upper / lower case | Employees!H:I | `=UPPER(B2)` · `=LOWER(B2)` |
| Range from a sheet name | Analysis | `=SUM(INDIRECT("'"&B11&"'!H2:H21"))` |
| Dynamic range for trends | Analysis | `=SUM(OFFSET(Monthly!B2,MATCH(B6,Monthly!A2:A5,0)-1,0,1,B12))` |
| Age from date of birth | Employees!J | `=DATEDIF(E2,TODAY(),"Y")` |
| Days between two dates | Analysis | `=MAX(Sales!A2:A21)-MIN(Sales!A2:A21)` |

### 5️⃣ Math & Filtering

| Task | Sheet | Formula |
|---|---|---|
| Round, ceiling, floor | Analysis | `=ROUND(x,2)` · `=CEILING(x,5)` · `=FLOOR(x,5)` |
| Top-performing students | TopPerformers | `=FILTER(Students!B2:F13,Students!F2:F13>B1,"No match")` |

---

## 🚀 Quick Start

1. **Open** `Fundamental_Booster.xlsx` in Excel or Google Sheets.
2. **Go to** the `Analysis` sheet.
3. **Edit** any yellow cell (region, product, student ID, salesperson, month...).
4. **Watch** the results update. The formula text sits beside every result.
5. **Experiment** with `Sales!P2` (tax rate) and `TopPerformers!B1` (score threshold).

> 🎨 **Colour code:** 🟨 blue text on yellow = input · ⬛ black text = formula

---

## 📌 Assumptions

| Item | Rule |
|---|---|
| Data | All names, scores, prices, dates and salaries are sample data |
| Currency | Shown in `$`; change the number format for other currencies |
| Grades | A+ ≥ 90 · A ≥ 80 · B ≥ 70 · C ≥ 60 · D ≥ 50 · else F (on average) |
| Discounts | Unit price ≥ 500 → 15% · ≥ 100 → 10% · ≥ 50 → 5% · else 0% |
| Eligibility | Unit price ≥ 500 **or** quantity ≥ 10 |
| Tax | 18%, stored once in `Sales!P2` |
| Dates | Age and tenure use `TODAY()` and change daily |

---

## 🔄 Compatibility

> [!NOTE]
> `XLOOKUP`, `XMATCH` and `FILTER` need **Excel 2021 or Microsoft 365**.
> To keep the workbook working everywhere, the live cells use `INDEX/MATCH` and a helper-column method that return identical results. The native 365 formulas are written next to them (`Analysis` rows 35–37 and `TopPerformers!E4`) so you can paste them in directly.

---

## 🛠 Skills Demonstrated

`Cell Referencing` · `Conditional Logic` · `Conditional Aggregation` · `Lookup Functions` · `Text Manipulation` · `Dynamic Ranges` · `Date Arithmetic` · `Financial Rounding` · `Data Filtering` · `Documentation`

---

<div align="center">

**Made with 📊 for PR. 1 — Fundamental Booster**

*Replace this line with your name, course and date.*

</div>
