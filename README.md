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

<img width="256" height="111" alt="image" src="https://github.com/user-attachments/assets/a7f3e0b2-e58e-4b58-b912-6fdb57b1612e" />

Key Finding

Adults aged 35–64 generated the highest profit: $93,496. This represents approximately 57% of the $163,508 total profit, making this age group the strongest contributor among the three segments shown.

📊 Analysis 2 — Sales Volume by Product and County

A second PivotTable compares sales volume across counties and three product categories:

<img width="453" height="187" alt="image" src="https://github.com/user-attachments/assets/d5873762-6614-4367-8d95-f675a464c583" />

The PivotTable was rearranged by switching the row and column fields, demonstrating how Excel can provide different analytical views without changing the underlying data.

🔄 Analysis 3 — Sales Volume Classification

A SWITCH formula was used to convert numerical sales volumes into easy-to-read performance categories.

=SWITCH(TRUE,C2>600,"high",C2>=300,"medium","low")

<img width="873" height="307" alt="image" src="https://github.com/user-attachments/assets/7e13ef2d-f9c2-49df-990f-63eee20f46b4" />

This technique makes numerical results easier to interpret and can be useful for dashboards, KPI reporting, and performance monitoring.

💡 Business Insights

The analysis highlights several useful findings:

Adults aged 35–64 are the most profitable customer segment in the age-group summary.

Laptops lead overall product sales volume, suggesting stronger demand than printers and smartphones in the exercise dataset.

Cornwall and Essex are the strongest counties by total sales volume, each reaching 1,100 units.

PivotTables allow the same dataset to be viewed from multiple perspectives quickly.

Categorising values as High, Medium, and Low makes performance easier for stakeholders to interpret.


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
