# E-Commerce Sales & Profitability Analysis (Excel Project)

##  Project Overview
This Excel analytics project evaluates e-commerce transaction performance, regional revenue distribution, category profitability, and shipping modes using order-level transaction data.

The raw dataset contains detailed order records tracking customer segments, product categories, shipping details, discounts, sales revenue, and profit margins.

---

## Excel Architecture & Technical Approach

### 1. Advanced Data Cleaning & Power Query ETL
- Automated Data Transformation:** Leveraged **Power Query** to extract, normalize, and clean raw transaction records.
- **Data Standardization:** Converted date strings into structured standard date attributes (`Order Date`, `Ship Date`) and extracted calendar dimensions (`Year`, `Month`).
- **Anomalies & Missing Values:** Addressed data discrepancies, removed duplicate order rows, and created custom calculated columns for profit margin percentages.

### 2. Advanced Formulas & Calculated Metrics
- **Dynamic Metrics:** Applied core Excel functions (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `XLOOKUP`, `IF/AND`) to calculate aggregate totals and performance flags.
- **Profitability Logic:** Formulated custom column calculations to evaluate net profit margins across product lines and regional customer segments.

### 3. Pivot Tables & Interactive Dashboard
- **Dimensional Slicing:** Constructed multiple **Pivot Tables** to analyze sales and profit distribution across Product Categories (*Furniture*, *Office Supplies*, *Technology*), Regions (*South*, *West*, *Central*, *East*), and Customer Segments.
- **Interactive Visualizations:** Built an Excel Dashboard utilizing **Pivot Charts** paired with dynamic **Slicers** (Timeline, Region, Category) for cross-filtering and scenario analysis.

---

## Key Business Insights
- **Top Product Categories:** Evaluated net profit margins across Furniture, Office Supplies, and Technology to identify high-margin SKUs versus discount-heavy items.
- **Regional Performance:** Mapped revenue distribution across regional markets to pinpoint high-volume sales hubs and fulfillment bottlenecks.
- **Shipping Mode Efficiency:** Analyzed order counts and shipping timelines across delivery classes (Standard Class, Second Class, First Class, Same Day).

---

##  Repository Contents
- `Ecommerce-Sales-Project1(1)cc-updt.xlsx` - Main Excel workbook containing raw data (`DATA!`), Power Query transformations, Pivot Tables, and Dashboard layout.

