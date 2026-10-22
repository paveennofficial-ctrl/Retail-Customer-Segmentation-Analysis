# Retail-Customer-Segmentation-Analysis
# Retail Sales & Customer Analysis

## 📌 Project Overview

The **Retail Sales & Customer Analysis** project uses Microsoft Excel to transform retail transaction data into meaningful business insights. It focuses on sales performance, product performance, customer purchasing activity, country-wise revenue, and sales trends over time.

The project combines data cleaning, Excel formulas, PivotTables, PivotCharts, KPI calculations, and an interactive dashboard to support data-driven decision-making.

## 🎯 Project Objectives

- Calculate total sales revenue, order volume, quantity sold, customer count, and average order value.
- Identify top- and bottom-performing products based on recorded sales revenue.
- Compare sales performance across countries.
- Analyze monthly and quarterly sales trends.
- Explore customer purchasing activity and patterns.
- Identify data-quality issues, including missing values, duplicates, invalid values, and negative quantities.
- Build an interactive Excel dashboard using KPI cards, charts, slicers, and a date Timeline.

## 📊 Dataset Overview

- **Data source:** Kaggle, according to the project report.
- **Tool used:** Microsoft Excel.
- **Cleaned dataset:** 96,458 records and 15 columns, according to the project report.
- **Time period:** December 2010 – April 2011.
- **Geographic coverage:** 30 countries.
- **Distinct invoices:** 5,824, as reported.
- **Unique customer IDs:** Approximately 2,200, as reported.
- **Unique product codes:** Approximately 2,900, as reported.

### Main Dataset Columns

| Column | Description |
|---|---|
| Invoice | Invoice or transaction reference |
| StockCode | Product identifier |
| Description | Product description |
| Quantity | Units recorded for the transaction |
| InvoiceDate | Transaction date and time |
| Price | Price per unit |
| Customer ID | Customer identifier |
| Country | Customer's country |
| Sales Revenue | Transaction-line revenue, generally Quantity × Price |
| Quality Status | Indicates records that are valid or require investigation |
| Year | Year extracted from the invoice date |
| Quarter | Quarter derived from the invoice date |
| Month | Month name |
| Month Number | Numeric month used for chronological sorting |
| WeekDays | Day of the week |

*Dataset figures are based on the project documentation. Reconcile the workbook's final KPI totals before publishing definitive results.*

## 🛠️ Tools and Technologies

- Microsoft Excel
- Excel formulas and structured references
- Data cleaning and validation
- Sorting and filtering
- Excel Tables
- PivotTables and PivotCharts
- KPI cards
- Slicers and Timeline filters
- Dashboard design and data visualization

### Excel Functions Used

| Function | Purpose |
|---|---|
| `SUM` | Calculate total recorded sales revenue |
| `SUMIF` | Calculate revenue for a selected country or category |
| `SUMIFS` | Calculate revenue using multiple conditions |
| `AVERAGE` | Calculate average transaction-line revenue |
| `AVERAGEIF` | Calculate an average for a selected group |
| `AVERAGEIFS` | Calculate an average using multiple conditions |
| `COUNTIF` | Count rows matching one condition |
| `COUNTIFS` | Count rows matching multiple conditions |
| `IF` / `IFS` | Classify records using specified rules |
| `XLOOKUP` | Retrieve product information using a product code |
| `EOMONTH` | Calculate a month-end date |

**KPI note:** Order volume should use the distinct invoice count rather than the number of transaction rows. Average order value is calculated as revenue divided by distinct orders, using a consistent revenue and returns policy.

## 🧹 Data Cleaning and Preparation

The project report documents these preparation activities:

1. **Missing-value review:** Identify missing information and assess its effect before deciding how to handle records.
2. **Duplicate removal:** Remove confirmed duplicate rows while preserving legitimate repeated purchases.
3. **Data-type conversion:** Check date, quantity, price, and identifier formats.
4. **Invalid-value checks:** Review unexpected quantities, prices, dates, and identifiers.
5. **Text standardization:** Check inconsistent spacing, capitalization, and category labels.
6. **Date-field creation:** Derive Year, Quarter, Month, Month Number, and WeekDays.
7. **Revenue calculation:** Calculate transaction-line revenue using Quantity × Price.

### Data Quality Consideration

Negative quantities may represent returns, cancellations, or adjustments. They should be investigated rather than automatically deleted or treated as ordinary sales. Revenue and KPI totals should be reconciled with the cleaned records before final reporting.

## 📈 Analysis and Dashboard

The project includes the following analysis views:

- **Monthly Sales Revenue:** A line chart to review changes in revenue over time.
- **Top 10 Products:** A horizontal bar chart comparing the products with the highest recorded sales revenue.
- **Bottom 10 Products:** A view of products with the lowest recorded revenue under the selected ranking method.
- **Revenue by Country:** A chart comparing recorded revenue across countries.
- **Quarterly Sales Revenue:** A column chart comparing revenue across quarters.
- **KPI Cards:** Summary measures such as total revenue, distinct orders, unique customers, quantity, and average order value.
- **Interactive Dashboard:** Slicers and a date Timeline allow users to filter connected PivotTables and charts.

## 💡 Key Insights and Business Applications

The analysis is designed to help users:

- Identify periods with higher or lower recorded revenue.
- Recognize products contributing most to sales and investigate low-revenue products.
- Compare country-level revenue and customer activity.
- Explore customer purchasing patterns and potential retention opportunities.
- Improve inventory planning, market evaluation, and sales reporting.
- Identify data-quality issues that could affect business decisions.

Specific findings, such as the leading country, top-performing product, and peak sales month, should be confirmed from the validated Excel workbook rather than assumed.

## ⚠️ Project Limitations

- The dataset covers December 2010 to April 2011, so some quarters may be incomplete.
- Missing customer IDs can limit customer-level analysis.
- Negative-quantity records require a clearly documented returns policy.
- Recorded revenue alone does not establish profit or product profitability.

## 📁 Project Files

Upload these files to your GitHub repository and adjust the names below if necessary:

- `README.md` — Project documentation.
- `Retail customer project F.xlsx` — Excel workbook containing the analysis and dashboard.
- `Retail Sales Customer Analysis PPT(1).pptx` — Project presentation.
- `Retail Customer Segmentation.docx` — Detailed project report.

## 🚀 How to Explore the Project

1. Download or clone the GitHub repository.
2. Open the Excel workbook in Microsoft Excel.
3. Review the cleaned dataset and supporting analysis sheets.
4. Open the dashboard and use the available slicers or Timeline filters.
5. Review the PowerPoint presentation and project report for methodology and context.

Some features, including slicers, Timelines, and PivotTables, work best in the desktop version of Microsoft Excel.

## 📝 Conclusion

This project demonstrates how Microsoft Excel can be used for retail data cleaning, formula-based analysis, summarization, visualization, and interactive reporting. By analyzing sales, products, countries, customers, and time periods, the project provides a structured view of retail performance.

Careful handling of returns, consistent KPI definitions, and reconciliation of totals are essential for producing reliable business insights.

## 👨‍💻 Author

**Paveen Kumar N**

**Project:** Retail Sales & Customer Analysis  
**Tools:** Microsoft Excel | Data Cleaning | Data Analysis | Data Visualization | Dashboard Reporting

---

*Developed to demonstrate practical data-analysis and reporting skills using Microsoft Excel.*
