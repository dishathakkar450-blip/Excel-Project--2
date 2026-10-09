<div align="center">

📊 Excel Sales Analysis Project

Turning Raw Sales Data into Business Insights

Excel • Formulas • Customer Analysis • What-If Analysis • Dashboard

</div>

👋 About This Project

This project uses Microsoft Excel to explore a sample sales dataset and present useful business insights. The workbook includes customer analysis, monthly trends, descriptive statistics, regression, discount scenarios, summary tables, and dashboard-style reporting.

The goal is to demonstrate how Excel formulas and cell references can help transform order data into meaningful information.

🎯 Project Goals

Calculate sales, profit, cost, and profit margin.

Identify high-value customers and rank customers by sales.

Compare sales and profit across regions and product categories.

Track monthly sales and profit trends.

Test discount scenarios with What-If Analysis.

Explore relationships between sales and profit using regression.

Summarize findings through tables, charts, and KPIs.

🗂️ Workbook Guide

Sheet

What it contains

Cover

Project title and cover page

Contents

Navigation guide

Project Instructions

Tasks and requirements

Dataset

Sample order data and calculated fields

Customer Analysis

Customer-level sales, profit, order count, and ranking

High Value Customers

High-value customer threshold and lookup formulas

What-If Analysis

Discount scenarios and Goal Seek setup

Regression

Linear regression analysis

Descriptive Stats

Summary statistics, quartiles, outlier fences, and correlation

Monthly Trend

Monthly sales, profit, order count, and growth

Pivot Summary

Formula-based summaries by region and category

Charts

Data visualizations

Dashboard

KPI overview and key insights

🧾 Dataset at a Glance

The workbook contains 200 sample order records with fields such as:

Customer_ID · Customer_Name · Region · Product_Category · Sales · Quantity · Discount · Order_Date · Profit

Calculated fields include month start, gross sales before discount, estimated cost, profit margin, and a current timestamp.

Note: The dataset covers April 2024 to April 2025. The first and last months may be partial periods.

🧮 Tasks & Excel Formulas

1. Calculate Important Sales Metrics

Task

Example formula

Purpose

Month start

=DATE(YEAR(H2),MONTH(H2),1)

Finds the first date of the order's month

Gross sales

=E2/(1-G2)

Estimates sales before discount

Cost

=E2-I2

Calculates cost as sales minus profit

Profit margin

=I2/E2

Calculates profit as a share of sales

Current timestamp

=NOW()

Returns the current date and time

2. Analyze Customers

Task

Example formula

Purpose

Count orders

=COUNTIF(Dataset!$A$2:$A$201,A4)

Counts orders for a customer

Total sales

=SUMIFS(Dataset!$E$2:$E$201,Dataset!$A$2:$A$201,A4)

Adds the customer's sales

Total profit

=SUMIFS(Dataset!$I$2:$I$201,Dataset!$A$2:$A$201,A4)

Adds the customer's profit

Average discount

=AVERAGEIFS(Dataset!$G$2:$G$201,Dataset!$A$2:$A$201,A4)

Finds the average discount

Sales rank

=RANK(C4,$C$4:$C$33)

Ranks customers by sales

Top 10 label

=IF(F4<=10,"Top 10","")

Labels the top 10 customers

3. Find High-Value Customers

Task

Example formula

Purpose

Sales threshold

=PERCENTILE('Customer Analysis'!C4:C33,0.75)

Calculates the 75th percentile

Count high-value customers

=COUNTIF('Customer Analysis'!C4:C33,">="&B3)

Counts customers above the threshold

Return ranked sales

=LARGE('Customer Analysis'!$C$4:$C$33,A7)

Finds a sales value at a given rank

Find matching customer

=INDEX('Customer Analysis'!$A$4:$A$33,MATCH(B7,'Customer Analysis'!$C$4:$C$33,0))

Looks up the customer ID

4. Explore Discount Scenarios

Task

Example formula

Purpose

Total sales

=SUM(Dataset!$E$2:$E$201)

Calculates total sales

Total profit

=SUM(Dataset!$I$2:$I$201)

Calculates total profit

Average discount

=AVERAGE(Dataset!$G$2:$G$201)

Calculates the average discount

Scenario sales

=$B$7*(1-A18)

Estimates sales at a selected discount

Scenario profit

=B18-$B$8

Estimates profit with cost held fixed

Goal Seek: In Excel, select Data → What-If Analysis → Goal Seek and use the target/input cells described in the workbook.

Assumption: This is a simplified scenario model. It estimates gross sales from recorded sales and discount, keeps total cost fixed, and assumes quantity does not change.

5. Run Regression Analysis

Task

Formula

Purpose

R-squared

=RSQ(Dataset!$I$2:$I$201,Dataset!$E$2:$E$201)

Measures the fit of the linear model

Standard error

=STEYX(Dataset!$I$2:$I$201,Dataset!$E$2:$E$201)

Estimates residual error size

Intercept

=INTERCEPT(Dataset!$I$2:$I$201,Dataset!$E$2:$E$201)

Estimates profit when sales are zero

Sales coefficient

=SLOPE(Dataset!$I$2:$I$201,Dataset!$E$2:$E$201)

Estimates how profit changes with sales

Predicted profit

=B17+B18*B21

Calculates predicted profit

6. Calculate Descriptive Statistics

The workbook explores mean, median, mode, standard deviation, variance, range, quartiles, IQR, percentiles, outlier fences, confidence intervals, correlation, and histogram frequencies.

Examples:

=AVERAGE(Dataset!$E$2:$E$201)
=MEDIAN(Dataset!$E$2:$E$201)
=STDEV(Dataset!$E$2:$E$201)
=CORREL(Dataset!$E$2:$E$201,Dataset!$I$2:$I$201)

7. Review Monthly Trends

Metric

Example formula

Monthly sales

=SUMIFS(Dataset!$E$2:$E$201,Dataset!$K$2:$K$201,A4)

Monthly profit

=SUMIFS(Dataset!$I$2:$I$201,Dataset!$K$2:$K$201,A4)

Order count

=COUNTIFS(Dataset!$K$2:$K$201,A4)

Month-over-month growth

=IFERROR(B5/B4-1,"")

Profit margin

=C4/B4

8. Summarize by Region and Category

Example formula:

=SUMIFS(Dataset!$E$2:$E$201,Dataset!$C$2:$C$201,$A5,Dataset!$D$2:$D$201,B$4)

This formula totals sales when both the region and product category match. The $ signs help keep the correct row or column fixed when copying the formula.

9. Build Dashboard KPIs

Examples of dashboard metrics include total sales, total profit, profit margin, total orders, top customer, best-performing region, and best-performing category.

=SUM(Dataset!$E$2:$E$201)
=SUM(Dataset!$I$2:$I$201)
=COUNTA(Dataset!$A$2:$A$201)

🔒 Absolute & Mixed References

The dollar sign ($) keeps part or all of a cell reference fixed when a formula is copied.

Reference

Type

What stays fixed?

A4

Relative

Nothing; row and column can change

$A$4

Absolute

Column A and row 4

$A4

Mixed

Column A only

A$4

Mixed

Row 4 only

Example: Dataset!$E$2:$E$201 keeps the sales-data range fixed while copying a formula. In a summary table, $A5 fixes the region column and B$4 fixes the category header row.

🧰 Skills Demonstrated

Microsoft Excel

Formula building and cell referencing

SUMIFS, COUNTIFS, AVERAGEIFS

IF, IFERROR, INDEX, MATCH, LARGE, RANK

Date and time functions

Descriptive statistics and correlation

Linear regression

What-If Analysis and Goal Seek

Conditional formatting

Data summaries, charts, and dashboards

🚀 How to Use This Workbook

Download or clone this repository.

Open the .xlsx workbook in Microsoft Excel.

Start with the Contents sheet.

Review the Dataset and calculated columns.

Explore the analysis sheets and check the formulas.

Change the discount input in What-If Analysis to test scenarios.

Review the Charts and Dashboard sheets.

If formulas do not refresh, set Excel calculation to Automatic and recalculate.

⚠️ Notes

This is a sample dataset for learning and portfolio demonstration.

NOW() is volatile and can update whenever Excel recalculates.

Formula-based summary tables are used in the Pivot Summary sheet.

The discount model relies on simplified assumptions and is not a guarantee of real-world outcomes.

Some charts or Excel ToolPak outputs may need to be refreshed or recreated in desktop Excel.

👩‍💻 Author

Disha Thakkar
Course: Data Analysis

<div align="center">

Excel-based sales analysis project for a data analytics portfolio.

</div>
