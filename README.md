# My-Excel-Projects
🚲 Excel Bike Sales Analysis






📌 Project Overview

This project demonstrates practical Microsoft Excel data analysis using a Bike Sales dataset. The workbook uses Excel Tables, PivotTables, and conditional logic to transform transactional data into clear business insights.

The analysis focuses on customer profitability, product sales performance, geographic comparisons, and sales-volume classification.

🎯 Project Objectives

Analyse profit by customer age group.

Compare product sales volume across counties.

Use PivotTables to summarise and restructure data.

Apply Excel's SWITCH function to classify sales performance.

Present business findings in a simple, decision-friendly format.

🛠️ Skills & Tools

Microsoft Excel was used for:

Excel Tables

PivotTables

Sorting and filtering

Data aggregation

SWITCH formulas

Conditional business logic

Sales and profit analysis

Data interpretation

📊 Dataset

The Bike Sales dataset includes fields such as:

Category

Example Fields

Order information

Sales Order Number, Date, Day, Month, Year

Customer information

Customer Age, Age Group, Customer Gender

Geography

Country, State

Product information

Product Category, Sub-Category, Product Description

Sales metrics

Order Quantity, Unit Cost, Unit Price, Profit

📈 Analysis 1 — Profit by Age Group

A PivotTable was created to calculate total profit for each customer age group.

Age Group

Sum of Profit

Adults (35–64)

$93,496

Young Adults (25–34)

$53,962

Youth (<25)

$16,050

Grand Total

$163,508

Key Finding

Adults aged 35–64 generated the highest profit: $93,496. This represents approximately 57% of the $163,508 total profit, making this age group the strongest contributor among the three segments shown.

📊 Analysis 2 — Sales Volume by Product and County

A second PivotTable compares sales volume across counties and three product categories:

Laptops

Printers

Smartphones

The analysis produced a total sales volume of 5,200 units.

Product Totals

Product

Sales Volume

Laptops

2,450

Printers

1,500

Smartphones

1,250

Grand Total

5,200

County Totals

County

Sales Volume

Cornwall

1,100

Essex

1,100

Greater Manchester

1,000

Lancashire

750

Yorkshire

700

Durham

550

Grand Total

5,200

Key Findings

Laptops were the highest-volume product, accounting for 2,450 units.

Cornwall and Essex jointly recorded the highest county totals, with 1,100 units each.

The PivotTable was rearranged by switching the row and column fields, demonstrating how Excel can provide different analytical views without changing the underlying data.

🔄 Analysis 3 — Sales Volume Classification

A SWITCH formula was used to convert numerical sales volumes into easy-to-read performance categories.

=SWITCH(TRUE,C2>600,"high",C2>=300,"medium","low")

Classification Rules

Sales Volume

Classification

Greater than 600

High

300 to 600

Medium

Below 300

Low

This technique makes numerical results easier to interpret and can be useful for dashboards, KPI reporting, and performance monitoring.

💡 Business Insights

The analysis highlights several useful findings:

Adults aged 35–64 are the most profitable customer segment in the age-group summary.

Laptops lead overall product sales volume, suggesting stronger demand than printers and smartphones in the exercise dataset.

Cornwall and Essex are the strongest counties by total sales volume, each reaching 1,100 units.

PivotTables allow the same dataset to be viewed from multiple perspectives quickly.

Categorising values as High, Medium, and Low makes performance easier for stakeholders to interpret.

📸 Project Screenshots

Add your screenshots to an images folder in this GitHub repository, then use the examples below to display them in the README.

images/
├── bike-sales-data.png
├── profit-by-age-group.png
├── pivot-product-by-county.png
├── pivot-county-by-product.png
└── switch-classification.png

Then add images using:

![Bike Sales Dataset](images/bike-sales-data.png)
![Profit by Age Group PivotTable](images/profit-by-age-group.png)
![Product by County PivotTable](images/pivot-product-by-county.png)
![County by Product PivotTable](images/pivot-county-by-product.png)
![SWITCH Classification](images/switch-classification.png)

📂 Suggested Repository Structure

excel-bike-sales-analysis/
├── README.md
├── Day_3_Task_1_Bike_Sales.xlsx
└── images/
    ├── bike-sales-data.png
    ├── profit-by-age-group.png
    ├── pivot-product-by-county.png
    ├── pivot-county-by-product.png
    └── switch-classification.png

🚀 What This Project Demonstrates

This project demonstrates my ability to:

Work with structured datasets in Excel.

Build and modify PivotTables.

Summarise sales and profitability metrics.

Apply conditional logic with Excel formulas.

Identify meaningful patterns in business data.

Communicate analytical findings clearly.

Turn raw spreadsheet data into actionable insights.

👤 Author

Shadedemon89

This project forms part of my data analysis portfolio and demonstrates practical Excel skills applicable to junior Data Analyst and Data Technician roles.
