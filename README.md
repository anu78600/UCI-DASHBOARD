# 🛒 UCI E-Commerce Business Analysis Dashboard

An end-to-end data analytics project focused on understanding e-commerce sales performance, customer behavior, product performance, regional trends, and business risks using **Excel Power Query, Power BI, DAX, and SQL**.

## 📌 Project Overview

The objective of this project is to transform raw e-commerce data into meaningful business insights. I cleaned and prepared the data, developed an interactive Power BI dashboard, and used SQL to investigate important business questions.

The project follows a business-focused analytical workflow, from data preparation to visualization and SQL-based analysis.

##  Project Workflow

**Raw Data → Data Cleaning in Excel Power Query → Power BI Dashboard → SQL Business Analysis → Business Insights**

| Stage                    | Tool              | Purpose                                 |
| ------------------------ | ----------------- | --------------------------------------- |
| 1. Raw Data              | CSV               | Identify data quality issues            |
| 2. Data Cleaning         | Excel Power Query | Clean, transform, and prepare the data  |
| 3. Dashboard Development | Power BI          | Visualize KPIs and business performance |
| 4. Business Analysis     | PostgreSQL        | Answer business questions using SQL     |
| 5. Documentation         | GitHub            | Present the project and analysis        |

##  Data Cleaning & Preparation

I used **Excel Power Query** to clean and transform the raw dataset before building the dashboard.

### Cleaning steps performed

* **Date and Time:** Extracted time from `Order_Date` into a separate `Order_Time` column and converted the date using the English (United States) locale.
* **Region:** Corrected inconsistent values using Find and Replace, identified extra spaces using a custom column, and removed them with the Trim transformation.
* **Payment Method:** Applied Trim and replaced blank values with `Unknown`.
* **Customer Rating:** Replaced null values with the average rating of `4.1`.
* **Quantity:** Corrected inconsistent quantity values.
* **Total Sales:** Created a custom column by multiplying `Unit_Price` by `Quantity`.
* **Product Names:** Standardized inconsistent capitalization, shortened names, and extra spaces.
* **Sales Channel:** Standardized text capitalization and removed extra spaces.
* **Discount:** Corrected an invalid discount value of `110` to `10`.
* **Customer ID:** Replaced blank customer IDs with `Unknown`.

### Additional data-quality considerations

The raw practice dataset includes issues such as missing values, inconsistent text, invalid values, and duplicate records. These issues provide opportunities to practice data-quality checks and validation.

##  Power BI Dashboard

The dashboard provides an overview of e-commerce business performance through KPI cards, charts, and interactive slicers.

### Key Performance Indicators (KPIs)

* **Total Sales:** Overall sales value.
* **Total Orders:** Number of distinct orders.
* **Unique Customers:** Distinct identifiable customers, excluding `Unknown` IDs.
* **Average Order Value (AOV):** Average sales value per order.
* **Average Rating:** Average customer rating.
* **Return Rate:** Percentage of orders marked as returned.
* **Revenue at Risk %:** Percentage of sales value associated with returned orders.

### Dashboard Visualizations

* Sales by Product Category
* Monthly Sales Trends
* Sales by Region
* Total Orders and Return Orders by Region
* Customer Rating and Delivery Days
* Sales by Payment Method

### Interactive Filters

* Sales Channel
* Product Name
* Year

These filters allow users to explore sales performance across different products, channels, regions, and periods.

##  DAX Measures

The following measures support the dashboard calculations. The table name used in the formulas is `Ecommerce_Business_Analysis_Messy_Practice`.

### 1. Total Sales

```dax
Total Sales =
SUM(
    Ecommerce_Business_Analysis_Messy_Practice[Total_Sales]
)
```

**Purpose:** Calculates the sum of sales values.

### 2. Total Orders

```dax
Total orders =
DISTINCTCOUNT(
    Ecommerce_Business_Analysis_Messy_Practice[Order_ID]
)
```

**Purpose:** Counts each distinct order once.

### 3. Unique Customers

```dax
Unique Customer =
CALCULATE(
    DISTINCTCOUNT(
        Ecommerce_Business_Analysis_Messy_Practice[Customer_ID]
    ),
    Ecommerce_Business_Analysis_Messy_Practice[Customer_ID]
        <> "Unknown"
)
```

**Purpose:** Counts distinct customer IDs while excluding missing IDs represented by `Unknown`.

### 4. Average Order Value (AOV)

```dax
AoV =
DIVIDE(
    [Total Sales],
    [Total orders]
)
```

**Purpose:** Calculates the average sales value per distinct order.

### 5. Returned Orders

```dax
Returned Orders =
CALCULATE(
    DISTINCTCOUNT(
        Ecommerce_Business_Analysis_Messy_Practice[Order_ID]
    ),
    Ecommerce_Business_Analysis_Messy_Practice[Return_Flag] = "Yes"
)
```

**Purpose:** Counts distinct orders marked as returned.

### 6. Return Rate

```dax
Return Rate =
DIVIDE(
    [Returned Orders],
    [Total orders],
    0
)
```

**Purpose:** Calculates the percentage of orders marked as returned.

Format this measure as a percentage in Power BI.

### 7. Returned Sales

```dax
Returned Sales =
CALCULATE(
    [Total Sales],
    Ecommerce_Business_Analysis_Messy_Practice[Return_Flag] = "Yes"
)
```

**Purpose:** Calculates the sales value associated with rows marked as returned.

### 8. Revenue at Risk %

```dax
Revenue at Risk % =
DIVIDE(
    [Returned Sales],
    [Total Sales],
    0
)
```

**Purpose:** Measures the proportion of recorded sales value associated with returned items or orders.

*Note: This is a proxy for sales associated with returns, not necessarily the final financial loss. The result depends on how returns and sales are recorded in the dataset.*

##  SQL Business Analysis

SQL is used to investigate business performance independently of the Power BI dashboard.

The following are example PostgreSQL queries for the project's business questions. They assume the cleaned data has been imported into a table named `ecommerce_sales`.

### 1. What is the total sales value?

```sql
SELECT
    SUM(total_sales) AS total_sales
FROM ecommerce_sales;
```

### 2. How many distinct orders were placed?

```sql
SELECT
    COUNT(DISTINCT order_id) AS total_orders
FROM ecommerce_sales;
```

### 3. How many identifiable unique customers are there?

```sql
SELECT
    COUNT(DISTINCT customer_id) AS unique_customers
FROM ecommerce_sales
WHERE customer_id IS NOT NULL
  AND TRIM(customer_id) <> ''
  AND LOWER(TRIM(customer_id)) <> 'unknown';
```

### 4. Which product categories generate the highest sales?

```sql
SELECT
    product_category,
    SUM(total_sales) AS total_sales
FROM ecommerce_sales
GROUP BY product_category
ORDER BY total_sales DESC;
```

### 5. Which are the top five products by sales?

```sql
SELECT
    product_name,
    SUM(total_sales) AS total_sales
FROM ecommerce_sales
GROUP BY product_name
ORDER BY total_sales DESC
LIMIT 5;
```

### 6. Which region has the highest return rate?

```sql
SELECT
    region,
    COUNT(DISTINCT order_id) AS total_orders,
    COUNT(DISTINCT order_id) FILTER (
        WHERE LOWER(TRIM(return_flag)) = 'yes'
    ) AS returned_orders,
    ROUND(
        100.0 * COUNT(DISTINCT order_id) FILTER (
            WHERE LOWER(TRIM(return_flag)) = 'yes'
        ) / NULLIF(COUNT(DISTINCT order_id), 0),
        2
    ) AS return_rate_pct
FROM ecommerce_sales
GROUP BY region
ORDER BY return_rate_pct DESC;
```

*Assumption: `return_flag` represents whether an order was returned. If an order can contain multiple rows with different return statuses, validate the order-level return logic before interpreting the result.*

### Further Business Questions

* Which sales channel generates the highest revenue?
* Which customer segment has the highest average order value?
* Which category has the highest return rate?
* What percentage of total sales is associated with returned orders?
* Which month generates the highest sales?
* How do sales change month over month?
* Which products rank in the top three within each category?
* Which regions have high sales but also high return rates?

These questions provide opportunities to practice aggregation, filtering, grouping, conditional calculations, subqueries, CTEs, and window functions.

##  Business Value

This project demonstrates how data analytics can help a business:

* Identify its strongest-performing product categories.
* Compare sales performance across regions and channels.
* Monitor customer ratings and delivery performance.
* Identify product categories and regions with return-related issues.
* Understand sales trends and prioritize further investigation.

The findings can help management make more informed decisions about product performance, customer experience, and operational improvements.

##  Tools & Technologies

* **Excel Power Query:** Data cleaning and transformation.
* **Power BI:** Interactive dashboard and data visualization.
* **DAX:** KPI and business metric calculations.
* **PostgreSQL:** SQL-based business analysis.
* **GitHub:** Project documentation and portfolio presentation.

##  Dataset Note

The dataset used for this project is a custom-generated e-commerce practice dataset designed with a realistic business structure and intentional data-quality issues. It is not presented as actual company transaction data.

##  Key Learning Outcomes

* Practiced cleaning and transforming messy data with Power Query.
* Built an interactive dashboard using Power BI.
* Created DAX measures for sales, orders, customers, AOV, and returns.
* Practiced translating business problems into SQL queries.
* Connected data preparation, visualization, and business analysis into one workflow.

**Project Goal:** Turn raw data into reliable analysis that supports better business decisions.
