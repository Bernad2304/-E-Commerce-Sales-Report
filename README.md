<div align="center">

# 🛒 E-Commerce Sales Performance Dashboard

### Excel-based sales analytics dashboard covering Oct 2023 – Oct 2025

</div>

<br/>

## 📌 Overview

An interactive Excel dashboard built on a Sales Fact table joined with Customer, Product, and Store dimension tables. The workbook covers the full analyst workflow — data cleaning, imputation, transformation, descriptive statistics, pivot-based visualization, and a final dashboard — to surface sales, regional, and customer-loyalty trends.

<br/>

## 📊 Dashboard Snapshot

| Metric | Value |
|---|---|
| **Total Revenue** | $2,170,969.81 |
| **Total Orders** | 2,000 |
| **Avg Order Value** | $1,085.48 |
| **Unique Customers** | 492 |
| **Top Region** | North |
| **Top Category** | Sports |
| **Period Covered** | Oct 2023 – Oct 2025 |

<br/>

## 🔍 Key Insights

**1. Sales Performance**
- The **Sports** category contributes the largest share of total sales, generating **$540,882.82** in revenue.

**2. Region-wise Sales**
- The **North** region has the highest sales at **$451,616.75**.
- The **West** region has the lowest sales at **$270,676.41**.

**3. Loyalty Level & Payment Method**
- **Platinum** customers have the highest average transaction values.
- **PayPal** is the most commonly used payment method.
- Transactions are broadly distributed across the customer base rather than concentrated in a few accounts.

**4. Monthly Sales Distribution**
- **March** is the top-performing month, contributing **$210,704.52** in revenue.
- **Men's Casual Sneakers** is the highest-selling product.

<br/>

## 🛠️ Process & Methodology

### Data Cleaning
- Imported the raw dataset and structured it into named tables.
- Standardized `Name` and `City` fields using `TRIM(PROPER(text))`.
- Replaced inconsistent `"Unknown"` values in `ProductName` via Find & Replace.
- Formatted `Unit Price` and `Total Amount` as currency.

### Data Imputation
- **Cost** — filled missing values using category/sub-category averages:
- **Stock** — same logic applied via `AVERAGEIFS` on category/sub-category.
- **Quantity** — filled using `AVERAGEIF` grouped by `Product_ID`.
- **Loyalty Level** — missing values imputed as `"Uncategorized"`.

### Data Transformation
- Extracted month from order date: `=TEXT(date_value, "MMM")`.

### Lookups & Table Joins
- Used `VLOOKUP` to merge dimension tables, pulling in Customer Name, Product Category, Product Name, Store Region, and Loyalty Level.

### Conditional Logic & Formatting
- Applied `IF` formulas in the **Sales Performance** and **Order Range** columns to flag high-performing sales and bulk orders.
- Used conditional formatting to visually highlight these flagged rows.

### Descriptive Statistics
- Used `SUM`, `COUNT`, `AVERAGE`, `AVERAGEIF`, and `MAX` to calculate total sales, total orders, number of bulk orders, mean order value, and maximum sale amount.

### Visualization & Dashboard
- Built Pivot Tables and Pivot Charts from the cleaned dataset.
- Assembled all visuals into a single interactive Excel dashboard for stakeholder-ready reporting.

<br/>

## 🧰 Tools Used

`Excel` `Pivot Tables` `Pivot Charts` `VLOOKUP` `IF Formulas` `AVERAGEIFS` `Conditional Formatting` `Dashboard Design`

<br/>

## 📁 Repository Structure

<br/>

<div align="center">

📩 Questions or feedback? Reach out via [LinkedIn](https://linkedin.com/in/Bernad-Meckenzi-S)

</div>
