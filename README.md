# 📊 Northwind Sales Analysis – Excel & Access Integration

This project presents a comprehensive sales analysis based on the sample **Northwind database**, combining Microsoft Access and Excel for relational data modeling, pivot analysis, and custom DAX formula applications.

---

## 🧰 Tools & Technologies
- **Microsoft Excel** (Power Pivot, DAX, Pivot Tables, Charts)
- **Microsoft Access** (`.accdb` format)
- **Data Modeling & Relationships**
- **Custom DAX Functions** for transformation

---

## 📁 Repository Contents

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
Flags all products with missing or undefined `Quantity Per Unit` values, aiding data completeness checks.

### 💵 Products > $25 PT  
Displays a filtered view of all products with a unit price above $25—useful for analyzing premium inventory.

---

## 📈 Visual Dashboards

### 📊 Products Target > 150 (Bar Chart)  
Visualizes all products with target level orders exceeding 150. This helps identify high-demand inventory and aids inventory planning decisions.

### 🥜 EC Dried Fruits & Nuts Target % (Pie Chart)  
Provides a percentage breakdown of target order levels among dried fruits and nuts products. Built to quickly understand category distribution within a specific segment.

---

## 🎯 Learning Goals

- **Master relational data modeling** by connecting Access tables in Excel’s Data Model.  
- **Enhance analytical thinking** through the creation of meaningful pivot tables and custom filters.  
- **Apply DAX formulas** like `CONCATENATEX` for advanced data transformation and summarization.  
- **Develop effective visualizations** to communicate insights using bar and pie charts.  
- **Practice data cleaning** by identifying and resolving missing or irrelevant entries.  
- **Integrate tools** (Excel + Access) to simulate real-world business data analysis scenarios.


---

## 🤝 Contact

Got feedback or want to collaborate?

📬 [Open an issue](https://github.com/SeojunKim05/DataAnalysis-northwind-data-analysis/issues)  
💼 [LinkedIn](https://www.linkedin.com/in/seojun-kim-089b7b339)  
📫 Email: kseojun05@gmail.com

---

