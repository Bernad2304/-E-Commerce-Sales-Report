<div align="center">

<img src="https://capsule-render.vercel.app/api?type=waving&color=0:00A8E8,100:6C63FF&height=200&section=header&text=E-Commerce%20Sales%20Dashboard&fontSize=36&fontColor=ffffff&animation=fadeIn&fontAlignY=35&desc=Excel%20Analytics%20%7C%20Oct%202023%20-%20Oct%202025&descAlignY=55&descSize=18"/>

<img src="https://readme-typing-svg.demolab.com?font=Fira+Code&size=20&pause=1000&color=6C63FF&center=true&vCenter=true&width=650&lines=Cleaning+%2B+Imputing+%2B+Transforming+Raw+Sales+Data;Built+with+Pivot+Tables%2C+VLOOKUP+%26+Conditional+Logic;Turning+2%2C000+Orders+into+Decision-Ready+Insights" alt="Typing SVG" />

<br/>

<img src="https://img.shields.io/badge/-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white"/>
<img src="https://img.shields.io/badge/-Pivot%20Tables-217346?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-VLOOKUP-6C63FF?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-Dashboard-00A8E8?style=for-the-badge"/>

</div>

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:6C63FF,100:00A8E8&height=3&width=1000"/>

## 📌 Overview

An interactive Excel dashboard built on a Sales Fact table joined with Customer, Product, and Store dimension tables. The workbook covers the full analyst workflow — data cleaning, imputation, transformation, descriptive statistics, pivot-based visualization, and a final dashboard — to surface sales, regional, and customer-loyalty trends.

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:00A8E8,100:6C63FF&height=3&width=1000"/>

## 📊 Dashboard Snapshot

<div align="center">

<table>
<tr>
<td align="center"><b>💰 Total Revenue</b><br/>$2,170,969.81</td>
<td align="center"><b>📦 Total Orders</b><br/>2,000</td>
<td align="center"><b>🧾 Avg Order Value</b><br/>$1,085.48</td>
</tr>
<tr>
<td align="center"><b>👥 Unique Customers</b><br/>492</td>
<td align="center"><b>🌎 Top Region</b><br/>North</td>
<td align="center"><b>🏆 Top Category</b><br/>Sports</td>
</tr>
</table>

<img src="https://img.shields.io/badge/Period-Oct%202023%20--%20Oct%202025-6C63FF?style=for-the-badge&labelColor=black"/>

</div>

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:6C63FF,100:00A8E8&height=3&width=1000"/>

## 🔍 Key Insights

<table>
<tr>
<td width="50%" valign="top">

### 🏅 Sales Performance
The **Sports** category contributes the largest share of total sales, generating **$540,882.82** in revenue.

</td>
<td width="50%" valign="top">

### 🌎 Region-wise Sales
**North** leads with **$451,616.75** — **West** trails at **$270,676.41**, the widest regional gap in the dataset.

</td>
</tr>
<tr>
<td width="50%" valign="top">

### 💳 Loyalty & Payment
**Platinum** customers post the highest average transaction values. **PayPal** is the most-used payment method, and spend is broadly distributed across the customer base.

</td>
<td width="50%" valign="top">

### 📅 Monthly Distribution
**March** is the top-performing month at **$210,704.52**. **Men's Casual Sneakers** is the best-selling product overall.

</td>
</tr>
</table>

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:00A8E8,100:6C63FF&height=3&width=1000"/>

## 🛠️ Process & Methodology

<details>
<summary><b>🧹 Data Cleaning</b> — click to expand</summary>
<br/>

- Imported the raw dataset and structured it into named tables
- Standardized `Name` and `City` fields using `TRIM(PROPER(text))`
- Replaced inconsistent `"Unknown"` values in `ProductName` via Find & Replace
- Formatted `Unit Price` and `Total Amount` as currency

</details>

<details>
<summary><b>🧩 Data Imputation</b> — click to expand</summary>
<br/>

**Cost** — filled using category/sub-category averages:
```excel
=IF(ISBLANK(Cost), AVERAGEIFS(tbl_CategoryList[Cost],
   tbl_CategoryList[Sub_Category], [Sub_Category],
   tbl_CategoryList[Category], [Category]), Cost)
```

**Stock** — same `AVERAGEIFS` logic on category/sub-category.

**Quantity** — filled using `AVERAGEIF` grouped by `Product_ID`.

**Loyalty Level** — missing values imputed as `"Uncategorized"`.

</details>

<details>
<summary><b>🔄 Data Transformation</b> — click to expand</summary>
<br/>

Extracted month from order date:
```excel
=TEXT(date_value, "MMM")
```

</details>

<details>
<summary><b>🔗 Lookups & Table Joins</b> — click to expand</summary>
<br/>

Used `VLOOKUP` to merge dimension tables, pulling in Customer Name, Product Category, Product Name, Store Region, and Loyalty Level.

</details>

<details>
<summary><b>🚦 Conditional Logic & Formatting</b> — click to expand</summary>
<br/>

Applied `IF` formulas in the **Sales Performance** and **Order Range** columns to flag high-performing sales and bulk orders, then used conditional formatting to visually highlight them.

</details>

<details>
<summary><b>📈 Descriptive Statistics</b> — click to expand</summary>
<br/>

Used `SUM`, `COUNT`, `AVERAGE`, `AVERAGEIF`, and `MAX` to calculate total sales, total orders, number of bulk orders, mean order value, and maximum sale amount.

</details>

<details>
<summary><b>📊 Visualization & Dashboard</b> — click to expand</summary>
<br/>

Built Pivot Tables and Pivot Charts from the cleaned dataset, then assembled all visuals into a single interactive Excel dashboard for stakeholder-ready reporting.

</details>

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:6C63FF,100:00A8E8&height=3&width=1000"/>

## 🧰 Tools Used

<div align="center">
<img src="https://img.shields.io/badge/-Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white"/>
<img src="https://img.shields.io/badge/-Pivot%20Tables-217346?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-Pivot%20Charts-217346?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-VLOOKUP-6C63FF?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-IF%20Formulas-6C63FF?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-AVERAGEIFS-6C63FF?style=for-the-badge"/>
<img src="https://img.shields.io/badge/-Conditional%20Formatting-00A8E8?style=for-the-badge"/>
</div>

<img src="https://capsule-render.vercel.app/api?type=rect&color=0:00A8E8,100:6C63FF&height=3&width=1000"/>

## 📁 Repository Structure
<img src="https://capsule-render.vercel.app/api?type=waving&color=0:6C63FF,100:00A8E8&height=120&section=footer"/>

</div>
