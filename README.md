# Excel Dashboard Portfolio

A collection of interactive Excel dashboards built to demonstrate data cleaning, formula-driven analysis, and dashboard design.

## 📁 Repository Contents

| File | Description |
|---|---|
| `Bike_Buyers_Dashboard.xlsx` | Customer demographics vs. bike purchase behavior (Pivot Tables & Slicers) |
| `Salary_Dashboard.xlsx` | Interactive global data-job salary explorer (Dynamic Arrays & Map Chart) |

---

## 1. Bike Buyers Analysis Dashboard

An interactive dashboard analyzing customer demographics and purchasing behavior to identify patterns in bike purchases.

### Overview

Explores a customer dataset (1,026 records) to understand which demographic and behavioral factors are associated with bike purchases — including income, gender, age, and commute distance.

### Dataset

13 fields per customer: Marital Status, Gender, Income, Children, Education, Occupation, Home Owner, Cars, Commute Distance, Region, Age, and Purchased Bike (Yes/No target).

### Workbook Structure

| Sheet | Purpose |
|---|---|
| `bike_buyers` | Raw, unmodified source data |
| `Working Sheet` | Cleaned data plus an engineered **Age Brackets** column (Adolescent / Middle Age / Old) |
| `Pivot Table` | Summary tables (average income and customer counts by gender, age bracket, and commute distance, split by purchase status) |
| `Dashboard` | Final interactive dashboard combining charts and slicers |

### Dashboard Features

- **Average Income by Gender & Purchase Status** — bar chart comparing income across buyers vs. non-buyers
- **Number of Customers by Commute Distance** — line chart showing how commute length relates to purchase behavior
- **Number of Customers by Age Bracket** — line chart showing purchase trends across age groups
- **Interactive Slicers** — filter the entire dashboard by Marital Status, Education, and Region

### Key Insights

- Bike buyers tend to have higher average income than non-buyers, across both genders
- Customers with the shortest commutes (0–1 miles) account for the largest share of both buyers and non-buyers
- Middle-aged customers make up the largest segment of both bike buyers and non-buyers

### Excel Skills Used

- **Pivot Tables** — built from the raw dataset to summarize income and customer counts across multiple dimensions
- **PivotCharts** — bar and line charts linked directly to pivot table data
- **Slicers** — interactive filtering across Marital Status, Education, and Region, connected to the pivot tables
- **Formula-based Feature Engineering** — nested `IF()` formulas to bucket ages into brackets
- **Data Structuring** — separated raw data, working/staging data, summary tables, and final dashboard into distinct sheets
- **Custom Number Formatting** — currency formatting (₱) applied to income figures
- **Dashboard Design** — single-page layout combining multiple visuals for at-a-glance analysis

### 🚀 How to Use

1. Download `Bike_Buyers_Dashboard.xlsx`
2. Open in Microsoft Excel
3. Go to the **Dashboard** sheet
4. Use the slicers to filter by Marital Status, Education, or Region and explore how the charts update

### 🧰 Tools

- Microsoft Excel (Pivot Tables, PivotCharts, Slicers, Formulas)

---

## 2. Data Job Salary Dashboard

An interactive, dropdown-driven dashboard for exploring global data-job postings and salaries.

### Overview

Analyzes a dataset of **32,672 job postings** to let a user pick a **Job Title**, **Country**, and **Employment Type** from dropdowns and instantly see the median salary, how that role compares to others, salary by employment type, salary by country on a map, and top hiring platforms — all without pivot tables, powered entirely by dynamic array formulas.

### Dataset

The `Data` sheet (Excel Table `jobs`) contains one row per job posting with fields including: job title, job location, job posting site, schedule type, remote status, posted date, degree/health-insurance mentions, country, salary rate/amount, company name, and required skills.

### Workbook Structure

| Sheet | Purpose |
|---|---|
| `Salary Dashboard` | Final dashboard: title/country/type selectors, charts, and a salary-by-country map |
| `Data` | Raw job postings data, stored as a structured Excel Table (`jobs`) |
| `Data Validation` | Dynamically generated, sorted unique lists (job titles, countries, schedule types) that feed the dashboard's dropdowns |
| `Jobs` | Median salary per job title, calculated live from the current Country/Type selection |
| `Country` | Median salary per country, calculated live from the current Job Title/Type selection |
| `Type` | Median salary per employment type, calculated live from the current Job Title/Country selection |
| `Platform` | Job posting counts per platform (e.g. LinkedIn, Indeed), for the current selection |

### Dashboard Features

- **Dropdown Selectors** — Job Title, Country, and Employment Type, sourced from live, sorted, deduplicated lists
- **Median Salary Comparison Charts** — bar charts by job title and by employment type, with the currently selected item highlighted in a different color
- **Filled Map Chart** — median salary by country shown geographically
- **Platform Breakdown** — count of postings by job site for the current filter combination
- **Fully Dynamic** — every table and chart recalculates instantly from the 32K-row dataset when a dropdown changes, with no pivot tables or manual refresh required

### Key Insights

- Median salary varies meaningfully by job title, with senior/specialized roles commanding a clear premium
- Country selection has a large effect on median salary for the same job title
- A small number of platforms account for the majority of postings for most job titles

### Excel Skills Used

- **Dynamic Array Formulas** — `UNIQUE()`, `SORT()`, `FILTER()`, and `ANCHORARRAY()` to build live, spill-based lookup lists and rankings
- **XLOOKUP** — used to pull the selected item's value out of the dynamically sorted arrays
- **Array-based MEDIAN/IF Formulas** — conditional median salary calculated directly over 32K+ rows based on three simultaneous filters
- **COUNTIFS** — platform-level counts filtered by all three dashboard selections at once
- **Excel Tables (Structured References)** — raw data stored as a Table (`jobs[...]`) so formulas stay readable and auto-expand
- **Named Ranges / Cell Names** — dropdown cells (`title`, `country`, `type`) named for use directly inside formulas
- **Data-Validation Dropdowns with Dynamic Sources** — dropdown lists driven by spilled array ranges rather than static lists
- **Filled Map Chart** — geographic visualization of median salary by country
- **Conditional Chart Highlighting** — split-series formulas (`IF` returning `NA()`) to color the selected bar differently from the rest
- **Large Dataset Handling** — formulas and charts built to stay responsive across ~32,700 rows

### 🚀 How to Use

1. Download `Salary_Dashboard.xlsx`
2. Open in Microsoft Excel (requires a recent version that supports dynamic arrays and `XLOOKUP`)
3. Go to the **Salary Dashboard** sheet
4. Use the Job Title, Country, and Employment Type dropdowns to filter and explore

### 🧰 Tools

- Microsoft Excel (Dynamic Arrays, XLOOKUP, Map Charts, Structured Tables, Formulas)
