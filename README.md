# PR. 1 – Fundamental Booster (Excel Formulas)

A hands-on Excel workbook that applies core spreadsheet functions to realistic student, sales and employee data. Every result is a live formula, so changing an input recalculates the whole workbook.

**File:** `Fundamental_Booster.xlsx`

---

## Objectives

- Use relative, absolute and mixed cell references correctly
- Classify and decide with `IF`, nested `IF`, `IF + AND` and `IF + OR`
- Summarise data with `COUNTIFS`, `SUMIFS` and `AVERAGEIFS`
- Retrieve data with `VLOOKUP`, `INDEX/MATCH`, `XLOOKUP` and `XMATCH`
- Clean and transform text with `LEFT`, `FIND`, `UPPER` and `LOWER`
- Build dynamic references with `INDIRECT` and `OFFSET`
- Work with dates (`DATEDIF`, `TODAY`) and maths (`ROUND`, `CEILING`, `FLOOR`)
- Return multiple values with `FILTER`

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Analysis** | Input cells (yellow) plus one formula result per task, with the formula text shown beside it |
| **Students** | Scores, average, nested-IF grade, `IF(AND)` check, helper column for filtering |
| **Products** | Product code, name and unit price (lookup table) |
| **Sales** | 20 transactions with VLOOKUP price, discount tiers, eligibility, net and tax-inclusive totals |
| **Employees** | Joining date, date of birth, salary, text functions, age and tenure |
| **Monthly** | Sales per salesperson per month (`SUMIFS`, mixed references) |
| **TopPerformers** | Dynamic list of students above a threshold (FILTER-style) |
| **References** | Absolute/relative/mixed reference demo and formatting samples |

---

## Task Coverage

| Topic | Task | Where | Key formula |
|---|---|---|---|
| References & formatting | Enter and format data; use `$A$1` vs `A1` | References, Sales!M | `=L2*(1+$P$2)` |
| IF / nested IF | Classify student grades | Students!G | `=IF(F2>=90,"A+",IF(F2>=80,"A",...))` |
| IF / nested IF | Discounts by price threshold | Sales!I | `=IF(F2>=500,0.15,IF(F2>=100,0.1,...))` |
| IF + AND | Math and Science both above 80 | Students!H | `=IF(AND(C2>80,D2>80),"Yes","No")` |
| IF + OR | Product discount eligibility | Sales!K | `=IF(OR(F2>=500,G2>=10),"Eligible","Not Eligible")` |
| COUNTIFS | Students scoring above 50 / 60 | Analysis | `=COUNTIFS(Students!C2:C13,">50")` |
| SUMIFS | Sales by region and product | Analysis | `=SUMIFS(Sales!H2:H21,Sales!C2:C21,B3,Sales!D2:D21,B4)` |
| AVERAGEIFS | Average of scores above 60 | Analysis | `=AVERAGEIFS(Students!F2:F13,Students!F2:F13,">60")` |
| VLOOKUP | Student name by ID; price by code | Analysis, Sales!F | `=VLOOKUP(B8,Products!A2:C6,3,FALSE)` |
| INDEX/MATCH | Sales by salesperson and month; dynamic employee field | Analysis | `=INDEX(Monthly!B2:D5,MATCH(B6,Monthly!A2:A5,0),MATCH(B7,Monthly!B1:D1,0))` |
| TEXT functions | First name, upper and lower case | Employees!G:I | `=LEFT(B2,FIND(" ",B2)-1)` |
| XLOOKUP | Salesperson performance; salary without sorting | Analysis | `=XLOOKUP(B9,Employees!A2:A11,Employees!F2:F11,"Not found")` |
| XMATCH | Position of a product in the sales list | Analysis | `=XMATCH(B4,Sales!D2:D21)` |
| INDIRECT | Reference a range from a sheet name in a cell | Analysis | `=SUM(INDIRECT("'"&B11&"'!H2:H21"))` |
| OFFSET | Dynamic range for sales trends | Analysis | `=SUM(OFFSET(Monthly!B2,MATCH(B6,Monthly!A2:A5,0)-1,0,1,B12))` |
| Date and time | Age from date of birth; days between dates | Employees!J:K, Analysis | `=DATEDIF(E2,TODAY(),"Y")` |
| Math | Rounding for financial values | Analysis | `=ROUND(x,2)`, `=CEILING(x,5)`, `=FLOOR(x,5)` |
| FILTER | Top-performing students | TopPerformers | `=FILTER(Students!B2:F13,Students!F2:F13>B1,"No match")` |

---

## How to Use

1. Open `Fundamental_Booster.xlsx` in Excel (desktop or 365) or Google Sheets.
2. Go to the **Analysis** sheet and edit the yellow input cells (region, product, student ID, salesperson, month, and so on).
3. Watch the results update. The third column shows the exact formula behind each result.
4. Change the tax rate on **Sales!P2** or the score threshold on **TopPerformers!B1** to see the dependent sheets recalculate.

**Colour convention:** blue text on yellow = input; black = formula.

---

## Assumptions

- All names, scores, prices, dates and salaries are sample data created for this exercise.
- Currency is shown in `$`; change the number format if you use another currency.
- Grade scale: A+ ≥ 90, A ≥ 80, B ≥ 70, C ≥ 60, D ≥ 50, otherwise F (applied to the average score).
- Discount tiers by unit price: ≥ 500 → 15%, ≥ 100 → 10%, ≥ 50 → 5%, otherwise 0%.
- A sale is discount eligible when unit price ≥ 500 **or** quantity ≥ 10.
- Tax rate is 18%, stored once in `Sales!P2` and used through an absolute reference.
- Age and tenure use `TODAY()`, so they change each day.

## Compatibility Note

`XLOOKUP`, `XMATCH` and `FILTER` require Excel 2021 / Microsoft 365. To keep the workbook working in every version, the live cells use `INDEX/MATCH` and a helper-column approach that return the same results. The native 365 formulas are written out next to them (Analysis rows 35–37 and TopPerformers!E4) so you can paste them in directly.

---

## Skills Demonstrated

Cell referencing · Conditional logic · Conditional aggregation · Lookup and reference functions · Text manipulation · Dynamic ranges · Date arithmetic · Financial rounding · Data filtering · Workbook documentation
