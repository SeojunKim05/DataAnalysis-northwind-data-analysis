# 📊 Northwind Sales Analysis – Excel & Access Integration

This project presents a comprehensive sales analysis based on the sample **Northwind database**, combining Microsoft Access and Excel for relational data modeling, pivot analysis, and custom DAX formula applications.

---

## 🧰 Tools & Technologies
- **Microsoft Excel** (Power Pivot, DAX, Pivot Tables, Charts)
- **Microsoft Access** (`.accdb` format)
- **Data Modeling & Relationships**
- **Custom DAX Functions** for transformation and aggregation

---

## 📁 Repository Files

- `Northwind Database.accdb`  
  Original sample database used as the external data source.

- `NorthwindAnalysis.xlsx`  
  The final Excel analysis file containing:
  - Cleaned, imported tables
  - Established data relationships
  - Analytical pivot tables and custom charts

---

## 🔗 Data Import & Modeling

I imported and cleaned data from **Access to Excel**, focusing on key tables:

- **Customers**
- **Orders**
- **Employees**
- **Products**

Unnecessary columns were removed to streamline analysis. I then created relationships:
- Between `Orders` and `Customers`
- Between `Orders` and `Employees`

This relationship is demonstrated on the 'Orders' table, EmployeeID and CustomerID associated with each order.
This enabled accurate, relationally-aware visualizations and calculations throughout the workbook.

---

## 📊 Pivot Tables & Visualizations

### 🧑‍💼 Purchasing Managers PT  
Identifies purchasing managers among customers using a filtered pivot table.  
To enhance clarity, I used `CONCATENATEX` in DAX to format customer names from separate first and last name columns.

### 🏠 Sales Reps – Redmond PT  
Filters employees by city (Redmond) and job title (Sales Representative), displaying relevant employee records via a custom pivot table.

### 💳 Credit Card Orders > $50 PT  
Highlights high-shipping-fee orders (shipping > $50) and displays the corresponding order ID and shipping city. This required targeted filtering and table joins through the data model.

### 🧃 Product Name – Category PT  
Organizes all products under their respective categories (e.g., *Beverages → Beer, Chai, etc.*), helping visualize product diversity across the company’s catalog.

### 🚫 Products Without Quantity Per Unit PT  
Flags all products with missing or undefined `QuantityPerUnit` values, aiding data completeness checks.

### 💵 Products > $25 PT  
Displays a filtered view of all products with a unit price above $25—useful for analyzing premium inventory.

---

## 📈 Visual Dashboards

### 📊 Products Target > 150 (Bar Chart)  
Visualizes all products with target level orders exceeding 150. This helps identify high-demand inventory and aids inventory planning decisions.

### 🥜 EC Dried Fruits & Nuts Target % (Pie Chart)  
Provides a percentage breakdown of target order levels among dried fruits and nuts products. Built to quickly understand category distribution within a specific segment.

---

## 🧠 Highlights of My Contributions

- Proactively **cleaned and structured** external Access data for analysis-ready use.
- Designed **relational data models** enabling complex multi-table analysis in Excel.
- Applied **custom DAX functions** to manipulate and unify data across dimensions (e.g., name formatting, conditional filtering).
- Built a series of **interactive pivot tables and visualizations** that surface actionable insights from the Northwind dataset.
- Ensured **data quality** by identifying and surfacing incomplete records.

---

## 📌 Summary

This project demonstrates how Excel can be leveraged as a powerful BI tool when combined with structured relational data and advanced analytical functions. It reflects my ability to:
- Work across systems (Access + Excel)
- Build data models from scratch
- Implement analytical logic
- Communicate findings visually and interactively

---

## 🔎 Next Steps

- Expand dashboard with KPIs (e.g., revenue by region, average shipping time)
- Automate data refresh using Power Query
- Migrate to Power BI for broader deployment

---

## 💬 Questions or Feedback?

Feel free to open an issue or connect with me if you're interested in collaboration or want to know more about this project.
