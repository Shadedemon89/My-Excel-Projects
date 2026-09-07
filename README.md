# My-Excel-Projects
# 📊 Excel Bike Sales & SWITCH Function Analysis

## 📌 Project Overview

This project demonstrates practical Microsoft Excel data analysis techniques using a Bike Sales dataset.

The workbook contains two main tasks:

1. **Bike Sales Analysis** – analysing sales volumes by product and county using PivotTables.
2. **SWITCH Function Task** – demonstrating the use of Excel's `SWITCH` function to create logic-based results from data.

The project was completed as part of my Excel/data analytics learning journey and demonstrates my ability to transform raw data into structured analysis and actionable insights.

---

## 🎯 Project Objectives

The main objectives of this project were to:

- Analyse bike-related sales data using Microsoft Excel.
- Summarise sales volumes by county and product.
- Use PivotTables to identify sales patterns.
- Compare product performance across different counties.
- Calculate overall sales totals.
- Practise the Excel `SWITCH` function.
- Develop practical spreadsheet and data-analysis skills.

---

## 🛠️ Tools & Technologies

- **Microsoft Excel**
- PivotTables
- Excel formulas
- `SWITCH` function
- Data summarisation
- Data analysis

---

## 📂 Workbook Structure

The workbook contains several worksheets used for the different stages of the analysis:

| Worksheet | Purpose |
|-----------|---------|
| `Sheet1` | Supporting worksheet / initial analysis |
| `Bike Sales` | Bike sales data and analysis |
| `Sheet2` | PivotTable analysis |
| `Sheet3` | Alternative PivotTable view |
| `SWITCH etc.` | SWITCH function exercise |

---

# 🚲 Task 1 – Bike Sales Analysis

The Bike Sales task uses PivotTables to analyse **Sales Volume** by:

- Product
- County

The products analysed are:

- Laptops
- Printers
- Smartphones

The counties included in the analysis are:

- Cornwall
- Durham
- Essex
- Greater Manchester
- Lancashire
- Yorkshire

---

## 📊 Sales Volume by Product

The PivotTable shows the following overall sales volumes:

| Product | Total Sales Volume |
|---------|-------------------:|
| Laptops | 2,450 |
| Printers | 1,500 |
| Smartphones | 1,250 |
| **Grand Total** | **5,200** |

### Key Finding

**Laptops were the highest-selling product**, with a total sales volume of **2,450**, representing almost half of the total recorded sales volume.

Printers generated **1,500** sales, while Smartphones generated **1,250**.

---

## 📍 Sales Volume by County

The analysis also summarises total sales volume by county:

| County | Total Sales Volume |
|--------|-------------------:|
| Cornwall | 1,100 |
| Essex | 1,100 |
| Greater Manchester | 1,000 |
| Lancashire | 750 |
| Yorkshire | 700 |
| Durham | 550 |
| **Grand Total** | **5,200** |

### Key Findings

- **Cornwall and Essex** recorded the highest total sales volumes at **1,100 each**.
- **Greater Manchester** followed with **1,000**.
- **Durham** recorded the lowest total sales volume at **550**.
- The combined sales volume across all counties was **5,200**.

---

## 🔎 Product Performance by County

The PivotTable allows product performance to be compared across individual counties.

| County | Laptops | Printers | Smartphones | Total |
|--------|--------:|---------:|------------:|------:|
| Cornwall | 700 | 400 | 0 | 1,100 |
| Essex | 0 | 800 | 300 | 1,100 |
| Greater Manchester | 400 | 0 | 600 | 1,000 |
| Lancashire | 600 | 0 | 150 | 750 |
| Yorkshire | 500 | 0 | 200 | 700 |
| Durham | 250 | 300 | 0 | 550 |
| **Total** | **2,450** | **1,500** | **1,250** | **5,200** |

### Observations

- Cornwall generated the majority of its sales through **Laptops**.
- Essex recorded particularly strong **Printer** sales.
- Greater Manchester was the strongest county for **Smartphones**.
- Lancashire's sales were primarily driven by **Laptops**.
- Yorkshire also showed strong Laptop sales.
- Durham generated sales from Laptops and Printers but no recorded Smartphone sales.

---

# 🔀 Task 2 – Excel SWITCH Function

The second task focuses on Excel's `SWITCH` function.

`SWITCH` is useful when a value needs to be compared against multiple possible conditions and a corresponding result returned.

A simplified example is:

```excel
=SWITCH(A2,
    1,"Low",
    2,"Medium",
    3,"High",
    "Unknown"
)
