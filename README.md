# 📊 Sales Analyzer — Excel Data Analysis Project

## 📌 Project Overview

**Sales Analyzer** is an Excel-based data analysis project designed to demonstrate practical skills in data cleaning, analysis, business intelligence, statistical analysis, what-if analysis, regression, customer analysis, and dashboard creation.

The workbook uses a formula-driven approach to transform sales data into meaningful business insights.

The project includes an **Executive Dashboard**, customer analysis, what-if analysis, regression analysis, descriptive statistics, monthly growth analysis, pivot-style summaries, and automated data validation checks.

---

## 🎯 Project Objectives

The main objectives of this project are to:

- Analyze sales and profit performance.
- Identify high-value and top customers.
- Compare sales across regions and product categories.
- Analyze monthly sales and profit growth.
- Perform statistical analysis on sales data.
- Understand the relationship between sales and profit.
- Perform what-if analysis for discount scenarios.
- Build an executive-level dashboard.
- Create automated integrity and error checks.
- Practice advanced Excel formulas and data-analysis techniques.

---

## 📂 Workbook Structure

The workbook contains the following sheets:

| Sheet | Purpose |
|---|---|
| **Dashboard** | Executive dashboard containing key sales and profit KPIs |
| **Dataset** | Main sales dataset used throughout the workbook |
| **Customer_Analysis** | Customer-level sales, profit, ranking, and segmentation analysis |
| **What_If** | Discount and profit scenario analysis |
| **Regression** | Linear regression analysis of Profit vs Sales |
| **Descriptive_Stats** | Statistical analysis and histogram calculations |
| **Monthly_Growth** | Monthly sales, profit, order, and growth analysis |
| **Pivot_Summary** | Region × Product Category sales summary |
| **Notes_Checks** | Task mapping, assumptions, and automated integrity checks |

---

# 📊 Dashboard

The **Dashboard** sheet provides an executive-level summary of the sales data.

### Key Performance Indicators

The dashboard includes:

- **Total Sales**
- **Total Profit**
- **Profit Margin**
- **Total Orders**
- **Average Sales per Order**
- **Average Profit per Order**
- **Customer Count**

The dashboard also provides automatically generated key insights based on the underlying dataset.

### Dashboard Features

- Dynamic date range
- Automatic KPI calculations
- Sales performance overview
- Profit performance overview
- Customer information
- Formula-driven insights
- Automatically refreshed timestamp

---

# 🗃️ Dataset

The **Dataset** sheet contains the main sales transaction data.

### Dataset Columns

| Column | Description |
|---|---|
| `Customer_ID` | Unique customer identifier |
| `Customer_Name` | Customer name |
| `Region` | Sales region |
| `Product_Category` | Product category |
| `Sales` | Sales amount |
| `Quantity` | Quantity sold |
| `Discount` | Discount applied |
| `Order_Date` | Date of order |
| `Profit` | Profit generated |
| `Timestamp` | Workbook refresh timestamp |

The dataset contains **200 sales orders** and is used as the source for all major calculations in the workbook.

---

# 👥 Customer Analysis

The **Customer_Analysis** sheet analyzes customer-level performance.

### Analysis Includes

- Total orders per customer
- Total sales per customer
- Total profit per customer
- Average order value
- Profit margin
- Customer sales ranking
- Customer segmentation
- Top 10 customers
- High-value customers

### Top Customer Analysis

Customers are ranked according to their total sales.

Conditional formatting is used to highlight the **Top 10 customers**.

### High-Value Customer Analysis

Customers are classified using a percentile-based threshold.

The workbook calculates:

- High-value sales threshold
- Number of high-value customers
- Customer segment
- Customer ranking

### Excel Functions Used

- `SUMIF`
- `COUNTIF`
- `PERCENTILE`
- `RANK`
- `INDEX`
- `MATCH`
- `IF`

---

# 🔄 What-If Analysis

The **What_If** sheet evaluates how changes in discount levels could affect sales and profit.

### Scenario Inputs

Users can change:

- Discount percentage change
- Expected volume uplift per percentage point of discount

The workbook then calculates the impact on:

- Average discount
- Net sales
- Profit
- Profit change
- Percentage change

### Sensitivity Analysis

A sensitivity table allows different discount scenarios to be compared at once.

This demonstrates how Excel can be used for **business scenario planning and decision-making**.

### Concepts Practiced

- What-if analysis
- Scenario modelling
- Sensitivity analysis
- `SUMPRODUCT`
- Conditional calculations
- Business assumptions

---

# 📈 Regression Analysis

The **Regression** sheet performs a linear regression analysis using:

> **Y = Profit**  
> **X = Sales**

The analysis is designed to resemble the output of:

**Excel Data Analysis ToolPak → Regression**

### Regression Statistics

The sheet calculates:

- Multiple R
- R Square
- Adjusted R Square
- Standard Error
- Number of observations

### ANOVA Analysis

The regression section also includes ANOVA-related calculations such as:

- Degrees of freedom
- Sum of Squares
- Mean Square
- F statistic
- Significance F

### Excel Functions Used

- `CORREL`
- `RSQ`
- `STEYX`
- Mathematical formulas

This analysis helps understand the relationship between **Sales and Profit**.

---

# 📊 Descriptive Statistics

The **Descriptive_Stats** sheet provides statistical analysis of:

- Sales
- Quantity
- Discount
- Profit

### Statistical Measures

The workbook calculates:

- Mean
- Standard Error
- Median
- Mode
- Standard Deviation
- Sample Variance
- Kurtosis
- Minimum
- Maximum
- Range
- Count

A histogram frequency analysis is also included.

### Excel Functions Used

- `AVERAGE`
- `STDEV`
- `MEDIAN`
- `MODE`
- `VAR`
- `KURT`
- `MIN`
- `MAX`
- `COUNT`

This section demonstrates how descriptive statistics can be used to understand the distribution and behavior of business data.

---

# 📅 Monthly Growth Analysis

The **Monthly_Growth** sheet analyzes sales performance month by month.

### Metrics

- Number of orders
- Monthly sales
- Monthly profit
- Sales change
- Sales growth percentage
- Profit growth percentage
- Average order value

### Growth Analysis

Month-over-month calculations are used to identify changes in business performance.

The workbook uses formulas to automatically calculate:

**Sales Change**

```excel
=Current Month Sales - Previous Month Sales
```

**Sales Growth %**

```excel
=(Current Month Sales - Previous Month Sales) / Previous Month Sales
```

**Average Order Value**

```excel
=Sales / Orders
```

### Excel Functions Used

- `COUNTIFS`
- `SUMIFS`
- `DATE`
- `YEAR`
- `MONTH`
- `EDATE`
- `IF`

---

# 📌 Pivot-Style Summary

The **Pivot_Summary** sheet provides a cross-tab analysis of:

> **Region × Product Category**

### Product Categories

- Books
- Clothing
- Electronics
- Furniture
- Office Supplies

### Regions

- Central
- East
- North
- South
- West

### Analysis Includes

- Sales by region
- Sales by product category
- Regional grand totals
- Product category totals
- Percentage share
- Regional ranking

The summary is created using `SUMIFS` formulas, allowing it to update automatically when the underlying dataset changes.

### Excel Functions Used

- `SUMIFS`
- `SUM`
- `RANK`

---

# ✅ Data Integrity & Error Checks

The **Notes_Checks** sheet contains automated validation checks to ensure that the workbook calculations are consistent.

### Integrity Checks Include

1. Pivot total sales matches dataset total sales.
2. Pivot total profit matches dataset total profit.
3. Monthly sales total matches dataset sales total.
4. Customer sales total matches dataset sales total.
5. Histogram frequencies match the number of orders.
6. A 0% what-if scenario matches baseline profit.
7. Regression uses all available orders.

Each check returns a status such as:

```text
OK
```

or

```text
CHECK
```

This helps verify that the workbook calculations are working correctly.

---

# 🧮 Excel Skills Demonstrated

This project demonstrates practical knowledge of:

### Basic & Intermediate Excel

- Formulas
- Cell references
- Conditional logic
- Number formatting
- Percentage calculations
- Date calculations

### Advanced Excel

- `SUMIFS`
- `COUNTIFS`
- `SUMIF`
- `COUNTIF`
- `INDEX`
- `MATCH`
- `RANK`
- `PERCENTILE`
- `SUMPRODUCT`
- `CORREL`
- `RSQ`
- `STEYX`
- Statistical functions
- Dynamic formulas

### Data Analysis

- Customer segmentation
- Ranking
- KPI analysis
- Sales analysis
- Profit analysis
- Growth analysis
- Statistical analysis
- Regression analysis
- Sensitivity analysis
- Cross-tab analysis

### Business Intelligence

- Executive dashboard
- KPI reporting
- Customer performance
- Regional analysis
- Product category analysis
- Business scenario analysis

---

# 📚 Key Learning Outcomes

After completing this project, the learner should be able to:

1. Build an Excel-based sales analysis project.
2. Create an executive dashboard using formulas.
3. Calculate sales and profit KPIs.
4. Analyze customer performance.
5. Identify top and high-value customers.
6. Perform discount-based what-if analysis.
7. Understand sales and profit relationships using regression.
8. Calculate descriptive statistics.
9. Analyze month-over-month growth.
10. Create region and product-category summaries.
11. Use advanced Excel functions for data analysis.
12. Build automated validation and error checks.
13. Convert raw business data into meaningful insights.

---

# 🛠️ Tools & Technologies

**Tool:**

- Microsoft Excel

**Techniques:**

- Data Analysis
- Business Intelligence
- Statistical Analysis
- Regression Analysis
- What-If Analysis
- Dashboard Development
- KPI Reporting
- Customer Segmentation
- Data Validation

---

# 🚀 How to Use This Project

1. Download or open the Excel workbook.
2. Start with the **Dashboard** sheet to view the overall business summary.
3. Explore the **Dataset** sheet to understand the source data.
4. Review **Customer_Analysis** for customer-level insights.
5. Modify the yellow inputs in **What_If** to test different discount scenarios.
6. Review the **Regression** sheet to understand Sales vs Profit.
7. Explore **Descriptive_Stats** for statistical analysis.
8. Check **Monthly_Growth** for month-over-month performance.
9. Review **Pivot_Summary** for regional and product-category analysis.
10. Use **Notes_Checks** to verify that the calculations reconcile correctly.

---

# 📁 Project File

```text
2. Analyzer (1).xlsx
```

---

# 📌 Project Highlights

✨ Formula-driven Excel dashboard  
✨ 200 sales transactions  
✨ Customer segmentation and ranking  
✨ Top 10 customer analysis  
✨ Discount sensitivity analysis  
✨ Sales vs Profit regression  
✨ Descriptive statistics  
✨ Monthly growth analysis  
✨ Region and product-category analysis  
✨ Automated integrity checks  

---

# 🎓 Project Type

**Excel Data Analysis & Business Intelligence Project**

This project is suitable for:

- Excel practice
- Data Analyst portfolio
- Business Analyst practice
- Excel interview preparation
- Learning advanced Excel functions
- Demonstrating data-analysis skills

---

## ⭐ Conclusion

This **Sales Analyzer** project demonstrates how Microsoft Excel can be used to transform raw sales data into useful business insights.

By combining dashboards, customer analysis, statistical analysis, regression, what-if scenarios, monthly growth analysis, and automated validation checks, the project provides a practical example of an end-to-end Excel data-analysis workflow.

⭐ **This project can be further extended with interactive slicers, Pivot Charts, advanced dashboards, Power Query, Power Pivot, and Power BI.**
