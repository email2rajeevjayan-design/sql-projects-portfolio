-- ============================================================
--   PROJECT 04: E-COMMERCE SALES ANALYTICS DASHBOARD
--   Tool: SQL Server Compatible
--   Dataset: 04_Ecommerce_Sales_Dataset.xlsx
-- ============================================================
--   SECTIONS:
--   1.  Database & Schema Setup
--   2.  Basic Exploration
--   3.  Sales & Revenue Analysis
--   4.  Product & Category Analysis
--   5.  Customer Segmentation Analysis
--   6.  Regional & Geographic Analysis
--   7.  Profitability & Discount Analysis
--   8.  Delivery, Returns & Operations
--   9.  Payment & Order Behaviour
--   10. Advanced: Window Functions
--   11. Advanced: CTEs & Subqueries
--   12. KPI Summary Dashboard Queries
-- ============================================================


-- ============================================================
-- SECTION 1: DATABASE & SCHEMA SETUP
-- ============================================================

CREATE DATABASE ecommerce_db;
USE ecommerce_db;

-- TABLE 1: Orders (main fact table)
DROP TABLE IF EXISTS orders;
CREATE TABLE orders (
    order_id            VARCHAR(12)     PRIMARY KEY,
    customer_id         VARCHAR(10)     NOT NULL,
    customer_name       VARCHAR(100),
    customer_email      VARCHAR(100),
    customer_segment    VARCHAR(20),
    order_date          DATE,
    ship_date           DATE,
    delivery_date       DATE,
    days_to_deliver     INT,
    region              VARCHAR(20),
    state               VARCHAR(50),
    city                VARCHAR(50),
    product_id          VARCHAR(10),
    product_name        VARCHAR(100),
    category            VARCHAR(50),
    sub_category        VARCHAR(50),
    brand               VARCHAR(50),
    quantity            INT,
    unit_price_inr      DECIMAL(10,2),
    discount_pct        DECIMAL(5,1),
    sales_amount_inr    DECIMAL(12,2),
    cogs_inr            DECIMAL(12,2),
    gross_profit_inr    DECIMAL(12,2),
    profit_margin_pct   DECIMAL(6,2),
    shipping_cost_inr   DECIMAL(8,2),
    payment_method      VARCHAR(30),
    order_status        VARCHAR(20),
    return_flag         VARCHAR(5),
    customer_rating     INT
);

-- TABLE 2: Product Master
DROP TABLE IF EXISTS product_master;
CREATE TABLE product_master (
    product_id          VARCHAR(10)     PRIMARY KEY,
    product_name        VARCHAR(100),
    category            VARCHAR(50),
    sub_category        VARCHAR(50),
    brand               VARCHAR(50),
    mrp_inr             DECIMAL(10,2),
    cost_price_inr      DECIMAL(10,2),
    stock_units         INT,
    avg_rating          DECIMAL(3,1),
    total_reviews       INT,
    launch_date         DATE,
    bestseller          VARCHAR(5)
);

-- TABLE 3: Customer Segments
DROP TABLE IF EXISTS customer_segments;
CREATE TABLE customer_segments (
    customer_id             VARCHAR(10)     PRIMARY KEY,
    customer_name           VARCHAR(100),
    segment                 VARCHAR(20),
    state                   VARCHAR(50),
    region                  VARCHAR(20),
    total_orders            INT,
    total_spend_inr         DECIMAL(12,2),
    avg_order_value         DECIMAL(10,2),
    last_order_date         DATE,
    days_since_last_order   INT,
    preferred_category      VARCHAR(50),
    preferred_payment       VARCHAR(30),
    customer_tier           VARCHAR(15),
    churn_risk              VARCHAR(10)
);

-- TABLE 4: Monthly KPIs
DROP TABLE IF EXISTS monthly_kpis;
CREATE TABLE monthly_kpis (
    year                INT,
    month               VARCHAR(10),
    category            VARCHAR(50),
    region              VARCHAR(20),
    total_orders        INT,
    revenue_inr         DECIMAL(15,2),
    cogs_inr            DECIMAL(15,2),
    gross_profit_inr    DECIMAL(15,2),
    profit_margin_pct   DECIMAL(6,2),
    avg_order_value     DECIMAL(10,2),
    returns             INT,
    return_rate_pct     DECIMAL(5,2),
    new_customers       INT,
    PRIMARY KEY (month, category, region)
);

-- IMPORT INSTRUCTIONS:
-- Export each Excel sheet to CSV, then use:
-- LOAD DATA INFILE '/path/to/Orders.csv'
-- INTO TABLE orders
-- FIELDS TERMINATED BY ',' ENCLOSED BY '"'
-- LINES TERMINATED BY '\n'
-- IGNORE 1 ROWS;
-- Repeat for product_master, customer_segments, monthly_kpis


-- ============================================================
-- SECTION 2: BASIC DATA EXPLORATION
-- ============================================================

-- 2.1 Preview all tables
SELECT * FROM orders;
SELECT * FROM product_master;
SELECT * FROM customer_segments;
SELECT * FROM monthly_kpis;

-- 2.2 Row counts
SELECT 'orders'             AS table_name, COUNT(*) AS total_rows FROM orders
UNION ALL
SELECT 'product_master',                   COUNT(*) FROM product_master
UNION ALL
SELECT 'customer_segments',                COUNT(*) FROM customer_segments
UNION ALL
SELECT 'monthly_kpis',                     COUNT(*) FROM monthly_kpis;

-- 2.3 Date range of orders
SELECT
    MIN(order_date)     AS first_order_date,
    MAX(order_date)     AS last_order_date,
    DATEDIFF(DAY, MIN(order_date), MAX(order_date)) AS span_days
FROM orders;

-- 2.4 Distinct categories and sub-categories
SELECT DISTINCT category, sub_category
FROM orders
ORDER BY category, sub_category;

-- 2.5 Unique customers, products, regions
SELECT
    COUNT(DISTINCT customer_id) AS unique_customers,
    COUNT(DISTINCT product_id)  AS unique_products,
    COUNT(DISTINCT region)      AS regions,
    COUNT(DISTINCT state)       AS states,
    COUNT(DISTINCT city)        AS cities
FROM orders;

-- 2.6 Order status distribution
SELECT
    order_status,
    COUNT(*)                                                        AS orders,
    ROUND(COUNT(*)*100.0/(SELECT COUNT(*) FROM orders), 2)         AS pct
FROM orders
GROUP BY order_status
ORDER BY orders DESC;

-- 2.7 Null / missing value check
SELECT
    SUM(CASE WHEN customer_rating IS NULL THEN 1 ELSE 0 END) AS null_ratings,
    SUM(CASE WHEN delivery_date   IS NULL THEN 1 ELSE 0 END) AS null_delivery,
    SUM(CASE WHEN shipping_cost_inr = 0   THEN 1 ELSE 0 END) AS free_shipping_orders
FROM orders;


-- ============================================================
-- SECTION 3: SALES & REVENUE ANALYSIS
-- ============================================================

-- 3.1 Overall revenue summary
SELECT
    COUNT(order_id)                         AS total_orders,
    ROUND(SUM(sales_amount_inr), 2)         AS total_revenue_inr,
    ROUND(AVG(sales_amount_inr), 2)         AS avg_order_value,
    ROUND(SUM(gross_profit_inr), 2)         AS total_gross_profit,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_profit_margin_pct,
    ROUND(SUM(shipping_cost_inr), 2)        AS total_shipping_cost
FROM orders
WHERE order_status NOT IN ('Cancelled');

-- 3.2 Monthly revenue trend
SELECT
    FORMAT(order_date,'yyyy-MM')        AS [month],
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(SUM(gross_profit_inr), 2)         AS gross_profit_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(sales_amount_inr), 2)         AS avg_order_value
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM')
ORDER BY FORMAT(order_date,'yyyy-MM');

-- 3.3 Quarterly revenue breakdown
SELECT
    YEAR(order_date)                                                    AS [year],
    CONCAT('Q', DATEPART(QUARTER,order_date))                                    AS quarter,
    COUNT(order_id)                                                     AS orders,
    ROUND(SUM(sales_amount_inr), 2)                                     AS revenue_inr,
    ROUND(SUM(gross_profit_inr), 2)                                     AS gross_profit_inr,
    ROUND(AVG(profit_margin_pct), 2)                                    AS avg_margin_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY YEAR(order_date), CONCAT('Q', DATEPART(QUARTER,order_date))
ORDER BY YEAR(order_date), CONCAT('Q', DATEPART(QUARTER,order_date));

-- 3.4 Year-over-year revenue comparison
SELECT
    YEAR(order_date)                                                    AS [year],
    ROUND(SUM(sales_amount_inr), 2)                                     AS revenue_inr,
    COUNT(order_id)                                                     AS total_orders,
    ROUND(AVG(sales_amount_inr), 2)                                     AS avg_order_value,
    ROUND(SUM(gross_profit_inr), 2)                                     AS gross_profit
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY YEAR(order_date)
ORDER BY YEAR(order_date);

-- 3.5 Revenue by day of week (which day drives most sales?)
SELECT
    DATENAME(WEEKDAY,order_date)                                                 AS day_of_week,
    DATEPART(WEEKDAY,order_date)                                               AS day_num,
    COUNT(order_id)                                                     AS orders,
    ROUND(SUM(sales_amount_inr), 2)                                     AS revenue_inr,
    ROUND(AVG(sales_amount_inr), 2)                                     AS avg_order_value
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY DATENAME(WEEKDAY,order_date), DATEPART(WEEKDAY,order_date)
ORDER BY DATEPART(WEEKDAY,order_date);

-- 3.6 Top revenue days
SELECT TOP 15
    order_date,
    COUNT(order_id)                     AS orders,
    ROUND(SUM(sales_amount_inr), 2)     AS daily_revenue
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY order_date
ORDER BY daily_revenue DESC;

-- 3.7 Revenue contribution by order size bucket
SELECT
    CASE
        WHEN sales_amount_inr < 500    THEN 'Micro (<500)'
        WHEN sales_amount_inr < 2000   THEN 'Small (500-1999)'
        WHEN sales_amount_inr < 10000  THEN 'Medium (2000-9999)'
        WHEN sales_amount_inr < 50000  THEN 'Large (10000-49999)'
        ELSE                                'Enterprise (50000+)'
    END                                 AS order_size_bucket,
    COUNT(*)                            AS orders,
    ROUND(SUM(sales_amount_inr), 2)     AS total_revenue,
    ROUND(AVG(profit_margin_pct), 2)    AS avg_margin_pct,
    ROUND(SUM(sales_amount_inr)*100.0 / (SELECT SUM(sales_amount_inr) FROM orders WHERE order_status!='Cancelled'), 2) AS revenue_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY
    CASE
        WHEN sales_amount_inr < 500    THEN 'Micro (<500)'
        WHEN sales_amount_inr < 2000   THEN 'Small (500-1999)'
        WHEN sales_amount_inr < 10000  THEN 'Medium (2000-9999)'
        WHEN sales_amount_inr < 50000  THEN 'Large (10000-49999)'
        ELSE                                'Enterprise (50000+)'
    END
ORDER BY total_revenue DESC;


-- ============================================================
-- SECTION 4: PRODUCT & CATEGORY ANALYSIS
-- ============================================================

-- 4.1 Revenue by category
SELECT
    category,
    COUNT(order_id)                         AS orders,
    SUM(quantity)                           AS units_sold,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(SUM(gross_profit_inr), 2)         AS gross_profit_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(sales_amount_inr), 2)         AS avg_order_value,
    ROUND(SUM(sales_amount_inr)*100.0/(SELECT SUM(sales_amount_inr) FROM orders WHERE order_status!='Cancelled'),2) AS revenue_share_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category
ORDER BY revenue_inr DESC;

-- 4.2 Sub-category deep dive
SELECT
    category,
    sub_category,
    COUNT(order_id)                         AS orders,
    SUM(quantity)                           AS units_sold,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(discount_pct), 2)             AS avg_discount_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category, sub_category
ORDER BY category, revenue_inr DESC;

-- 4.3 Top 15 best-selling products by revenue
SELECT TOP 15
    product_name,
    category,
    brand,
    COUNT(order_id)                         AS orders,
    SUM(quantity)                           AS units_sold,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(discount_pct), 2)             AS avg_discount_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY product_name, category, brand
ORDER BY revenue_inr DESC;

-- 4.4 Top products by units sold (volume leaders)
SELECT TOP 15
    product_name, category, brand,
    SUM(quantity)                           AS total_units_sold,
    COUNT(order_id)                         AS orders,
    ROUND(AVG(unit_price_inr), 2)           AS avg_price
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY product_name, category, brand
ORDER BY total_units_sold DESC;

-- 4.5 Brand performance
SELECT
    brand,
    COUNT(order_id)                         AS orders,
    SUM(quantity)                           AS units_sold,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(discount_pct), 2)             AS avg_discount_pct,
    ROUND(AVG(customer_rating), 2)          AS avg_customer_rating
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY brand
ORDER BY revenue_inr DESC;

-- 4.6 Product catalog: bestsellers vs non-bestsellers
SELECT
    bestseller,
    COUNT(*)                                AS products,
    ROUND(AVG(mrp_inr), 2)                 AS avg_mrp,
    ROUND(AVG(avg_rating), 2)              AS avg_rating,
    SUM(total_reviews)                      AS total_reviews,
    ROUND(AVG(stock_units), 1)             AS avg_stock
FROM product_master
GROUP BY bestseller;

-- 4.7 Low stock alert (products needing reorder)
SELECT
    product_id, product_name, category, sub_category,
    brand, stock_units, avg_rating, total_reviews
FROM product_master
WHERE stock_units < 50
ORDER BY stock_units ASC;

-- 4.8 Price range analysis
SELECT
    CASE
        WHEN mrp_inr < 500     THEN 'Budget (<500)'
        WHEN mrp_inr < 2000    THEN 'Mid-range (500-1999)'
        WHEN mrp_inr < 10000   THEN 'Premium (2000-9999)'
        WHEN mrp_inr < 50000   THEN 'Luxury (10000-49999)'
        ELSE                        'Ultra-Premium (50000+)'
    END                                     AS price_tier,
    COUNT(*)                                AS products,
    ROUND(AVG(avg_rating), 2)              AS avg_rating,
    SUM(CASE WHEN bestseller='Yes' THEN 1 ELSE 0 END) AS bestsellers
FROM product_master
GROUP BY
    CASE
        WHEN mrp_inr < 500     THEN 'Budget (<500)'
        WHEN mrp_inr < 2000    THEN 'Mid-range (500-1999)'
        WHEN mrp_inr < 10000   THEN 'Premium (2000-9999)'
        WHEN mrp_inr < 50000   THEN 'Luxury (10000-49999)'
        ELSE                        'Ultra-Premium (50000+)'
    END
ORDER BY products DESC;


-- ============================================================
-- SECTION 5: CUSTOMER SEGMENTATION ANALYSIS
-- ============================================================

-- 5.1 Customer tier distribution
SELECT
    customer_tier,
    COUNT(*)                                AS customers,
    ROUND(AVG(total_spend_inr), 2)         AS avg_spend,
    ROUND(AVG(total_orders), 1)            AS avg_orders,
    ROUND(AVG(avg_order_value), 2)         AS avg_order_value,
    ROUND(AVG(days_since_last_order), 1)   AS avg_recency_days
FROM customer_segments
GROUP BY customer_tier
ORDER BY avg_spend DESC;

-- 5.2 Churn risk analysis
SELECT
    churn_risk,
    COUNT(*)                                AS customers,
    ROUND(AVG(total_spend_inr), 2)         AS avg_lifetime_spend,
    ROUND(AVG(days_since_last_order), 1)   AS avg_days_inactive,
    ROUND(AVG(total_orders), 1)            AS avg_orders
FROM customer_segments
GROUP BY churn_risk
ORDER BY customers DESC;

-- 5.3 RFM-style analysis (Recency / Frequency / Monetary)
SELECT TOP 30
    customer_id,
    customer_name,
    customer_tier,
    days_since_last_order                   AS recency_days,
    total_orders                            AS frequency,
    ROUND(total_spend_inr, 2)               AS monetary_value,
    ROUND(avg_order_value, 2)               AS avg_order_value,
    churn_risk
FROM customer_segments
ORDER BY monetary_value DESC;

-- 5.4 Segment (B2C/B2B) performance
SELECT
    o.customer_segment                      AS segment,
    COUNT(DISTINCT o.customer_id)           AS customers,
    COUNT(o.order_id)                       AS total_orders,
    ROUND(SUM(o.sales_amount_inr), 2)       AS total_revenue,
    ROUND(AVG(o.sales_amount_inr), 2)       AS avg_order_value,
    ROUND(AVG(o.profit_margin_pct), 2)      AS avg_margin_pct
FROM orders o
WHERE order_status != 'Cancelled'
GROUP BY o.customer_segment
ORDER BY total_revenue DESC;

-- 5.5 High-value customers (Top 20% by revenue)
SELECT TOP 50
    customer_id,
    customer_name,
    customer_tier,
    total_orders,
    ROUND(total_spend_inr, 2)               AS lifetime_value,
    preferred_category,
    preferred_payment,
    churn_risk
FROM customer_segments
ORDER BY total_spend_inr DESC;

-- 5.6 Customer preferred categories
SELECT
    preferred_category,
    COUNT(*)                                AS customers,
    ROUND(AVG(total_spend_inr), 2)         AS avg_spend,
    ROUND(AVG(total_orders), 1)            AS avg_orders
FROM customer_segments
GROUP BY preferred_category
ORDER BY customers DESC;

-- 5.7 New vs returning customers trend by month
SELECT
    FORMAT(o.order_date,'yyyy-MM')       AS [month],
    COUNT(DISTINCT o.customer_id)           AS active_customers,
    COUNT(DISTINCT CASE WHEN cs.total_orders = 1 THEN o.customer_id END) AS new_customers,
    COUNT(DISTINCT CASE WHEN cs.total_orders > 1 THEN o.customer_id END) AS returning_customers
FROM orders o
LEFT JOIN customer_segments cs ON o.customer_id = cs.customer_id
WHERE o.order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM')
ORDER BY [month];


-- ============================================================
-- SECTION 6: REGIONAL & GEOGRAPHIC ANALYSIS
-- ============================================================

-- 6.1 Revenue by region
SELECT
    region,
    COUNT(order_id)                         AS orders,
    COUNT(DISTINCT customer_id)             AS customers,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(SUM(gross_profit_inr), 2)         AS gross_profit,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(sales_amount_inr), 2)         AS avg_order_value,
    ROUND(SUM(sales_amount_inr)*100.0/(SELECT SUM(sales_amount_inr) FROM orders WHERE order_status!='Cancelled'),2) AS revenue_share_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY region
ORDER BY revenue_inr DESC;

-- 6.2 Top 10 states by revenue
SELECT TOP 10
    state,
    region,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    COUNT(DISTINCT customer_id)             AS unique_customers
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY state, region
ORDER BY revenue_inr DESC;

-- 6.3 Top cities by revenue
SELECT TOP 15
    city, state, region,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(AVG(sales_amount_inr), 2)         AS avg_order_value
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY city, state, region
ORDER BY revenue_inr DESC;

-- 6.4 Category preference by region
SELECT
    region,
    category,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(COUNT(order_id)*100.0/SUM(COUNT(order_id)) OVER (PARTITION BY region),2) AS region_share_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY region, category
ORDER BY region, revenue_inr DESC;

-- 6.5 Regional shipping cost analysis
SELECT
    region,
    ROUND(AVG(shipping_cost_inr), 2)        AS avg_shipping_cost,
    ROUND(SUM(shipping_cost_inr), 2)        AS total_shipping_cost,
    SUM(CASE WHEN shipping_cost_inr = 0 THEN 1 ELSE 0 END) AS free_shipping_orders,
    ROUND(AVG(days_to_deliver), 1)          AS avg_delivery_days
FROM orders
GROUP BY region
ORDER BY avg_shipping_cost DESC;

-- 6.6 State-level growth comparison (2023 vs 2024)
SELECT TOP 15
    state,
    ROUND(SUM(CASE WHEN YEAR(order_date)=2023 THEN sales_amount_inr ELSE 0 END),2) AS revenue_2023,
    ROUND(SUM(CASE WHEN YEAR(order_date)=2024 THEN sales_amount_inr ELSE 0 END),2) AS revenue_2024,
    ROUND(
        (SUM(CASE WHEN YEAR(order_date)=2024 THEN sales_amount_inr ELSE 0 END) -
         SUM(CASE WHEN YEAR(order_date)=2023 THEN sales_amount_inr ELSE 0 END)) * 100.0 /
        NULLIF(SUM(CASE WHEN YEAR(order_date)=2023 THEN sales_amount_inr ELSE 0 END),0),
    2) AS yoy_growth_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY state
HAVING SUM(CASE WHEN YEAR(order_date)=2023 THEN sales_amount_inr ELSE 0 END) > 0
   AND SUM(CASE WHEN YEAR(order_date)=2024 THEN sales_amount_inr ELSE 0 END) > 0
ORDER BY yoy_growth_pct DESC;


-- ============================================================
-- SECTION 7: PROFITABILITY & DISCOUNT ANALYSIS
-- ============================================================

-- 7.1 Profit margin by category
SELECT
    category,
    ROUND(SUM(gross_profit_inr), 2)         AS total_profit,
    ROUND(SUM(sales_amount_inr), 2)         AS total_revenue,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(MIN(profit_margin_pct), 2)        AS min_margin,
    ROUND(MAX(profit_margin_pct), 2)        AS max_margin
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category
ORDER BY avg_margin_pct DESC;

-- 7.2 Discount impact on profitability
SELECT
    CASE
        WHEN discount_pct = 0   THEN 'No Discount'
        WHEN discount_pct <= 10 THEN 'Low (1-10%)'
        WHEN discount_pct <= 20 THEN 'Medium (11-20%)'
        WHEN discount_pct <= 30 THEN 'High (21-30%)'
        ELSE                         'Very High (30%+)'
    END                                     AS discount_tier,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(sales_amount_inr), 2)         AS avg_order_value,
    ROUND(AVG(customer_rating), 2)          AS avg_customer_rating
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY
    CASE
        WHEN discount_pct = 0   THEN 'No Discount'
        WHEN discount_pct <= 10 THEN 'Low (1-10%)'
        WHEN discount_pct <= 20 THEN 'Medium (11-20%)'
        WHEN discount_pct <= 30 THEN 'High (21-30%)'
        ELSE                         'Very High (30%+)'
    END
ORDER BY orders DESC;

-- 7.3 Category-wise discount strategy
SELECT
    category,
    ROUND(AVG(discount_pct), 2)             AS avg_discount_pct,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    COUNT(CASE WHEN discount_pct = 0  THEN 1 END) AS no_discount_orders,
    COUNT(CASE WHEN discount_pct >= 20 THEN 1 END) AS high_discount_orders,
    COUNT(order_id)                         AS total_orders
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category
ORDER BY avg_discount_pct DESC;

-- 7.4 Products with negative or very low margin (loss leaders)
SELECT TOP 15
    product_name, category, brand,
    ROUND(AVG(profit_margin_pct), 2)        AS avg_margin_pct,
    ROUND(AVG(discount_pct), 2)             AS avg_discount_pct,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr), 2)         AS revenue_inr
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY product_name, category, brand
HAVING ROUND(AVG(profit_margin_pct),2) < 15
ORDER BY avg_margin_pct ASC;

-- 7.5 Profit contribution analysis (80/20 rule)
SELECT
    category,
    sub_category,
    ROUND(SUM(gross_profit_inr), 2)         AS total_profit,
    ROUND(SUM(gross_profit_inr)*100.0/(SELECT SUM(gross_profit_inr) FROM orders WHERE order_status!='Cancelled'), 2) AS profit_share_pct,
    SUM(SUM(gross_profit_inr)*100.0/(SELECT SUM(gross_profit_inr) FROM orders WHERE order_status!='Cancelled'))
        OVER (ORDER BY SUM(gross_profit_inr) DESC ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS cumulative_profit_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category, sub_category
ORDER BY total_profit DESC;

-- 7.6 Revenue vs cost monthly trend
SELECT
    FORMAT(order_date,'yyyy-MM')         AS [month],
    ROUND(SUM(sales_amount_inr),2)          AS revenue_inr,
    ROUND(SUM(cogs_inr),2)                  AS total_cogs,
    ROUND(SUM(gross_profit_inr),2)          AS gross_profit,
    ROUND(SUM(shipping_cost_inr),2)         AS shipping_cost,
    ROUND(AVG(profit_margin_pct),2)         AS avg_margin_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM')
ORDER BY [month];


-- ============================================================
-- SECTION 8: DELIVERY, RETURNS & OPERATIONS
-- ============================================================

-- 8.1 Delivery performance overview
SELECT
    ROUND(AVG(days_to_deliver), 2)          AS avg_delivery_days,
    MIN(days_to_deliver)                    AS min_days,
    MAX(days_to_deliver)                    AS max_days,
    COUNT(CASE WHEN days_to_deliver <= 3 THEN 1 END) AS fast_delivery,
    COUNT(CASE WHEN days_to_deliver BETWEEN 4 AND 6 THEN 1 END) AS standard_delivery,
    COUNT(CASE WHEN days_to_deliver >= 7 THEN 1 END) AS slow_delivery,
    COUNT(*)                                AS total_orders
FROM orders
WHERE order_status = 'Delivered';

-- 8.2 Delivery time by region
SELECT
    region,
    ROUND(AVG(days_to_deliver), 2)          AS avg_delivery_days,
    MIN(days_to_deliver)                    AS min_days,
    MAX(days_to_deliver)                    AS max_days,
    COUNT(CASE WHEN days_to_deliver <= 3 THEN 1 END) AS fast_pct_orders
FROM orders
WHERE order_status = 'Delivered'
GROUP BY region
ORDER BY avg_delivery_days ASC;

-- 8.3 Return rate by category
SELECT
    category,
    COUNT(*)                                AS total_orders,
    SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END) AS returns,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate_pct,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN gross_profit_inr ELSE 0 END),2) AS profit_lost_to_returns
FROM orders
GROUP BY category
ORDER BY return_rate_pct DESC;

-- 8.4 Return rate by region
SELECT
    region,
    COUNT(*)                                AS total_orders,
    SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END) AS returns,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate_pct
FROM orders
GROUP BY region
ORDER BY return_rate_pct DESC;

-- 8.5 High return products (quality issues flag)
SELECT TOP 15
    product_name, category, brand,
    COUNT(*)                                AS orders,
    SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END) AS returned,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate_pct,
    ROUND(AVG(customer_rating),2)           AS avg_rating
FROM orders
GROUP BY product_name, category, brand
HAVING orders >= 10 AND return_rate_pct > 15
ORDER BY return_rate_pct DESC;

-- 8.6 Cancelled orders analysis
SELECT
    category,
    COUNT(CASE WHEN order_status='Cancelled' THEN 1 END) AS cancelled_orders,
    COUNT(*) AS total_orders,
    ROUND(COUNT(CASE WHEN order_status='Cancelled' THEN 1 END)*100.0/COUNT(*),2) AS cancel_rate_pct,
    ROUND(SUM(CASE WHEN order_status='Cancelled' THEN sales_amount_inr ELSE 0 END),2) AS revenue_lost
FROM orders
GROUP BY category
ORDER BY cancel_rate_pct DESC;

-- 8.7 Customer rating distribution
SELECT
    customer_rating,
    COUNT(*)                                AS orders,
    ROUND(COUNT(*)*100.0/SUM(COUNT(*)) OVER(),2) AS pct,
    ROUND(AVG(profit_margin_pct),2)        AS avg_margin,
    ROUND(AVG(days_to_deliver),2)           AS avg_delivery_days
FROM orders
WHERE customer_rating IS NOT NULL
GROUP BY customer_rating
ORDER BY customer_rating DESC;


-- ============================================================
-- SECTION 9: PAYMENT & ORDER BEHAVIOUR
-- ============================================================

-- 9.1 Payment method preference
SELECT
    payment_method,
    COUNT(order_id)                         AS orders,
    ROUND(COUNT(order_id)*100.0/(SELECT COUNT(*) FROM orders),2) AS usage_pct,
    ROUND(SUM(sales_amount_inr),2)          AS revenue_inr,
    ROUND(AVG(sales_amount_inr),2)          AS avg_order_value,
    ROUND(AVG(profit_margin_pct),2)         AS avg_margin_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY payment_method
ORDER BY orders DESC;

-- 9.2 Payment method by category
SELECT
    category,
    payment_method,
    COUNT(*)                                AS orders,
    ROUND(COUNT(*)*100.0/SUM(COUNT(*)) OVER (PARTITION BY category),2) AS pct_within_category
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category, payment_method
ORDER BY category, orders DESC;

-- 9.3 Payment method by customer segment
SELECT
    customer_segment,
    payment_method,
    COUNT(*)                                AS orders,
    ROUND(AVG(sales_amount_inr),2)          AS avg_order_value
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY customer_segment, payment_method
ORDER BY customer_segment, COUNT(*) DESC;

-- 9.4 COD (Cash on Delivery) vs digital payments
SELECT
    CASE WHEN payment_method = 'Cash on Delivery' THEN 'COD' ELSE 'Digital' END AS payment_type,
    COUNT(*)                                AS orders,
    ROUND(AVG(sales_amount_inr),2)          AS avg_order_value,
    ROUND(AVG(days_to_deliver),2)           AS avg_delivery_days,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate_pct,
    ROUND(AVG(customer_rating),2)           AS avg_rating
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY
    CASE WHEN payment_method = 'Cash on Delivery' THEN 'COD' ELSE 'Digital' END;

-- 9.5 Average order quantity by category
SELECT
    category,
    ROUND(AVG(quantity),2)                  AS avg_quantity_per_order,
    SUM(quantity)                           AS total_units,
    COUNT(order_id)                         AS orders
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category
ORDER BY avg_quantity_per_order DESC;


-- ============================================================
-- SECTION 10: WINDOW FUNCTIONS
-- ============================================================

-- 10.1 Rank categories by monthly revenue
SELECT
    FORMAT(order_date,'yyyy-MM')         AS [month],
    category,
    ROUND(SUM(sales_amount_inr),2)          AS revenue,
    RANK() OVER (PARTITION BY FORMAT(order_date,'yyyy-MM') ORDER BY SUM(sales_amount_inr) DESC) AS revenue_rank
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM'), category
ORDER BY FORMAT(order_date,'yyyy-MM'), revenue_rank;

-- 10.2 Running total revenue per category over months
SELECT
    FORMAT(order_date,'yyyy-MM')         AS [month],
    category,
    ROUND(SUM(sales_amount_inr),2)          AS monthly_revenue,
    ROUND(SUM(SUM(sales_amount_inr)) OVER (
        PARTITION BY category
        ORDER BY FORMAT(order_date,'yyyy-MM')
        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    ),2)                                    AS cumulative_revenue
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM'), category
ORDER BY category, FORMAT(order_date,'yyyy-MM');

-- 10.3 Month-over-month revenue growth per region
SELECT
    FORMAT(order_date,'yyyy-MM')         AS [month],
    region,
    ROUND(SUM(sales_amount_inr),2)          AS revenue,
    LAG(ROUND(SUM(sales_amount_inr),2)) OVER (
        PARTITION BY region
        ORDER BY FORMAT(order_date,'yyyy-MM')
    )                                       AS prev_month_revenue,
    ROUND(
        (SUM(sales_amount_inr) - LAG(SUM(sales_amount_inr)) OVER (
            PARTITION BY region ORDER BY FORMAT(order_date,'yyyy-MM')
        )) * 100.0 /
        NULLIF(LAG(SUM(sales_amount_inr)) OVER (
            PARTITION BY region ORDER BY FORMAT(order_date,'yyyy-MM')
        ),0)
    ,2)                                     AS mom_growth_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM'), region
ORDER BY region, FORMAT(order_date,'yyyy-MM');

-- 10.4 Customer revenue percentile (PERCENT_RANK)
SELECT
    customer_id, customer_name,
    ROUND(total_spend_inr,2)                AS lifetime_value,
    ROUND(PERCENT_RANK() OVER (ORDER BY total_spend_inr)*100,2) AS revenue_percentile,
    NTILE(5) OVER (ORDER BY total_spend_inr DESC) AS revenue_quintile
FROM customer_segments
ORDER BY lifetime_value DESC;

-- 10.5 Top product per category by revenue (ROW_NUMBER)
SELECT * FROM (
    SELECT
        category, product_name, brand,
        ROUND(SUM(sales_amount_inr),2)      AS revenue,
        ROUND(AVG(profit_margin_pct),2)     AS margin_pct,
        ROW_NUMBER() OVER (PARTITION BY category ORDER BY SUM(sales_amount_inr) DESC) AS rn
    FROM orders
    WHERE order_status != 'Cancelled'
    GROUP BY category, product_name, brand
) ranked
WHERE rn = 1
ORDER BY revenue DESC;

-- 10.6 3-month rolling average revenue per category
SELECT
    FORMAT(order_date,'yyyy-MM')         AS [month],
    category,
    ROUND(SUM(sales_amount_inr),2)          AS monthly_revenue,
    ROUND(AVG(SUM(sales_amount_inr)) OVER (
        PARTITION BY category
        ORDER BY FORMAT(order_date,'yyyy-MM')
        ROWS BETWEEN 2 PRECEDING AND CURRENT ROW
    ),2)                                    AS rolling_3m_avg
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM'), category
ORDER BY category, FORMAT(order_date,'yyyy-MM');

-- 10.7 Customer order ranking within their segment
SELECT
    customer_id, customer_name, segment,
    total_orders, ROUND(total_spend_inr,2)  AS total_spend,
    RANK() OVER (PARTITION BY segment ORDER BY total_spend_inr DESC) AS segment_rank,
    DENSE_RANK() OVER (ORDER BY total_spend_inr DESC)                AS overall_rank
FROM customer_segments
ORDER BY segment, segment_rank;

-- 10.8 Lead: anticipate next month orders for forecasting
SELECT
    FORMAT(order_date,'yyyy-MM')         AS [month],
    COUNT(order_id)                         AS orders,
    LEAD(COUNT(order_id)) OVER (ORDER BY FORMAT(order_date,'yyyy-MM')) AS next_month_orders_signal
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY FORMAT(order_date,'yyyy-MM')
ORDER BY [month];


-- ============================================================
-- SECTION 11: CTEs & SUBQUERIES
-- ============================================================

;
-- 11.1 CTE: Customer Lifetime Value (CLV) segmentation
WITH clv_base AS (
    SELECT
        o.customer_id,
        cs.customer_name,
        cs.customer_tier,
        COUNT(o.order_id)                   AS total_orders,
        ROUND(SUM(o.sales_amount_inr),2)    AS total_revenue,
        ROUND(SUM(o.gross_profit_inr),2)    AS total_profit,
        ROUND(AVG(o.sales_amount_inr),2)    AS avg_order_value,
        ROUND(AVG(o.profit_margin_pct),2)   AS avg_margin_pct,
        MAX(o.order_date)                   AS last_order_date,
        DATEDIFF(DAY, MAX(o.order_date), GETDATE()) AS days_since_last_order
    FROM orders o
    JOIN customer_segments cs ON o.customer_id = cs.customer_id
    WHERE o.order_status != 'Cancelled'
    GROUP BY o.customer_id, cs.customer_name, cs.customer_tier
),
clv_scored AS (
    SELECT *,
        ROUND(
            (total_profit * 0.5) +
            (total_orders * 500) -
            (days_since_last_order * 10),
        2) AS clv_score
    FROM clv_base
)
SELECT TOP 30
    customer_id, customer_name, customer_tier,
    total_orders, total_revenue, total_profit,
    avg_margin_pct, days_since_last_order, clv_score,
    RANK() OVER (ORDER BY clv_score DESC)  AS clv_rank
FROM clv_scored
ORDER BY clv_score DESC;

;
-- 11.2 CTE: Category health report
WITH category_stats AS (
    SELECT
        category,
        COUNT(order_id)                     AS orders,
        ROUND(SUM(sales_amount_inr),2)      AS revenue,
        ROUND(SUM(gross_profit_inr),2)      AS profit,
        ROUND(AVG(profit_margin_pct),2)     AS avg_margin,
        ROUND(AVG(discount_pct),2)          AS avg_discount,
        ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate,
        ROUND(AVG(customer_rating),2)       AS avg_rating
    FROM orders
    WHERE order_status != 'Cancelled'
    GROUP BY category
),
benchmarks AS (
    SELECT
        ROUND(AVG(avg_margin),2)            AS bench_margin,
        ROUND(AVG(avg_discount),2)          AS bench_discount,
        ROUND(AVG(return_rate),2)           AS bench_return_rate,
        ROUND(AVG(avg_rating),2)            AS bench_rating
    FROM category_stats
)
SELECT
    cs.category, cs.orders, cs.revenue, cs.profit,
    cs.avg_margin,
    CASE WHEN cs.avg_margin > b.bench_margin THEN 'Above Avg' ELSE 'Below Avg' END AS margin_vs_bench,
    cs.avg_discount,
    CASE WHEN cs.avg_discount > b.bench_discount THEN 'High Discount' ELSE 'Lean Discount' END AS discount_flag,
    cs.return_rate,
    CASE WHEN cs.return_rate > b.bench_return_rate THEN 'High Returns' ELSE 'Healthy' END AS return_flag,
    cs.avg_rating
FROM category_stats cs
CROSS JOIN benchmarks b
ORDER BY cs.revenue DESC;

;
-- 11.3 CTE: Multi-step sales funnel
WITH order_funnel AS (
    SELECT
        FORMAT(order_date,'yyyy-MM')     AS [month],
        COUNT(*)                            AS total_orders,
        COUNT(CASE WHEN order_status='Processing' THEN 1 END) AS processing,
        COUNT(CASE WHEN order_status='Shipped'    THEN 1 END) AS shipped,
        COUNT(CASE WHEN order_status='Delivered'  THEN 1 END) AS delivered,
        COUNT(CASE WHEN order_status='Cancelled'  THEN 1 END) AS cancelled,
        COUNT(CASE WHEN order_status='Returned'   THEN 1 END) AS returned
    FROM orders
    GROUP BY FORMAT(order_date,'yyyy-MM')
)
SELECT *,
    ROUND(delivered*100.0/total_orders,2)   AS delivery_rate_pct,
    ROUND(cancelled*100.0/total_orders,2)   AS cancellation_rate_pct,
    ROUND(returned*100.0/total_orders,2)    AS return_rate_pct
FROM order_funnel
ORDER BY [month];

;
-- 11.4 CTE: Regional product affinity
WITH region_category AS (
    SELECT
        region, category,
        ROUND(SUM(sales_amount_inr),2)      AS revenue,
        RANK() OVER (PARTITION BY region ORDER BY SUM(sales_amount_inr) DESC) AS cat_rank
    FROM orders
    WHERE order_status != 'Cancelled'
    GROUP BY region, category
)
SELECT region, category, revenue, cat_rank
FROM region_category
WHERE cat_rank <= 3
ORDER BY region, cat_rank;

;
-- 11.5 CTE: Discount elasticity (revenue impact of discounting)
WITH discount_buckets AS (
    SELECT
        category,
        CASE
            WHEN discount_pct = 0   THEN 'No Discount'
            WHEN discount_pct <= 10 THEN '1-10%'
            WHEN discount_pct <= 20 THEN '11-20%'
            ELSE                         '21%+'
        END                             AS discount_bucket,
        ROUND(AVG(sales_amount_inr),2)  AS avg_revenue,
        ROUND(AVG(profit_margin_pct),2) AS avg_margin,
        ROUND(AVG(quantity),2)          AS avg_qty,
        COUNT(*)                        AS orders
    FROM orders
    WHERE order_status != 'Cancelled'
    GROUP BY category,
        CASE
            WHEN discount_pct = 0   THEN 'No Discount'
            WHEN discount_pct <= 10 THEN '1-10%'
            WHEN discount_pct <= 20 THEN '11-20%'
            ELSE                         '21%+'
        END
)
SELECT *
FROM discount_buckets
ORDER BY category, avg_revenue DESC;

-- 11.6 Subquery: Products outperforming category average margin
SELECT
    o.product_name, o.category, o.brand,
    ROUND(AVG(o.profit_margin_pct),2)   AS product_margin,
    cat_avg.avg_cat_margin,
    ROUND(AVG(o.profit_margin_pct) - cat_avg.avg_cat_margin, 2) AS margin_advantage
FROM orders o
JOIN (
    SELECT TOP 20 category, ROUND(AVG(profit_margin_pct),2) AS avg_cat_margin
    FROM orders WHERE order_status != 'Cancelled'
    GROUP BY category
) cat_avg ON o.category = cat_avg.category
WHERE o.order_status != 'Cancelled'
GROUP BY o.product_name, o.category, o.brand, cat_avg.avg_cat_margin
HAVING ROUND(AVG(o.profit_margin_pct),2) > cat_avg.avg_cat_margin
ORDER BY margin_advantage DESC;

;
-- 11.7 CTE: Churn prediction signal (customers at risk)
WITH last_orders AS (
    SELECT
        customer_id,
        MAX(order_date)                     AS last_order,
        DATEDIFF(DAY, MAX(order_date), GETDATE()) AS days_inactive,
        COUNT(*)                            AS total_orders,
        ROUND(SUM(sales_amount_inr),2)      AS lifetime_value
    FROM orders
    WHERE order_status != 'Cancelled'
    GROUP BY customer_id
),
churn_flags AS (
    SELECT *,
        CASE
            WHEN days_inactive > 180 AND total_orders >= 5 THEN 'High Value - Churned'
            WHEN days_inactive > 180 AND total_orders < 5  THEN 'Low Value - Churned'
            WHEN days_inactive > 90  AND total_orders >= 5 THEN 'High Value - At Risk'
            WHEN days_inactive > 90  AND total_orders < 5  THEN 'Low Value - At Risk'
            ELSE 'Active'
        END AS churn_segment
    FROM last_orders
)
SELECT
    churn_segment,
    COUNT(*)                                AS customers,
    ROUND(AVG(lifetime_value),2)           AS avg_lifetime_value,
    ROUND(AVG(total_orders),1)             AS avg_orders,
    ROUND(AVG(days_inactive),1)            AS avg_days_inactive
FROM churn_flags
GROUP BY churn_segment
ORDER BY avg_lifetime_value DESC;


-- ============================================================
-- SECTION 12: KPI SUMMARY DASHBOARD QUERIES
-- ============================================================

-- 12.1 Executive KPI Scorecard
SELECT
    COUNT(DISTINCT order_id)                                AS total_orders,
    COUNT(DISTINCT customer_id)                             AS unique_customers,
    ROUND(SUM(sales_amount_inr),2)                          AS total_revenue_inr,
    ROUND(AVG(sales_amount_inr),2)                          AS avg_order_value_inr,
    ROUND(SUM(gross_profit_inr),2)                          AS total_gross_profit_inr,
    ROUND(AVG(profit_margin_pct),2)                         AS avg_profit_margin_pct,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS overall_return_rate_pct,
    ROUND(SUM(CASE WHEN order_status='Cancelled' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS cancellation_rate_pct,
    ROUND(AVG(days_to_deliver),2)                           AS avg_delivery_days,
    ROUND(AVG(customer_rating),2)                           AS avg_customer_rating
FROM orders;

-- 12.2 Category performance leaderboard
SELECT
    category,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr),2)          AS revenue_inr,
    ROUND(SUM(gross_profit_inr),2)          AS profit_inr,
    ROUND(AVG(profit_margin_pct),2)         AS margin_pct,
    ROUND(AVG(discount_pct),2)              AS avg_discount,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate,
    RANK() OVER (ORDER BY SUM(sales_amount_inr) DESC)       AS revenue_rank,
    RANK() OVER (ORDER BY AVG(profit_margin_pct) DESC)      AS margin_rank
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY category
ORDER BY revenue_rank;

-- 12.3 Regional performance summary
SELECT
    region,
    COUNT(DISTINCT customer_id)             AS customers,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr),2)          AS revenue_inr,
    ROUND(AVG(profit_margin_pct),2)         AS avg_margin,
    ROUND(AVG(days_to_deliver),2)           AS avg_delivery_days,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY region
ORDER BY revenue_inr DESC;

-- 12.4 Customer tier revenue contribution
SELECT
    cs.customer_tier,
    COUNT(DISTINCT o.customer_id)           AS customers,
    COUNT(o.order_id)                       AS orders,
    ROUND(SUM(o.sales_amount_inr),2)        AS revenue_inr,
    ROUND(AVG(o.sales_amount_inr),2)        AS avg_order_value,
    ROUND(SUM(o.sales_amount_inr)*100.0/(SELECT SUM(sales_amount_inr) FROM orders WHERE order_status!='Cancelled'),2) AS revenue_share_pct
FROM orders o
JOIN customer_segments cs ON o.customer_id = cs.customer_id
WHERE o.order_status != 'Cancelled'
GROUP BY cs.customer_tier
ORDER BY revenue_inr DESC;

-- 12.5 Year-over-year full business summary
SELECT
    YEAR(order_date)                        AS [year],
    COUNT(DISTINCT customer_id)             AS customers,
    COUNT(order_id)                         AS orders,
    ROUND(SUM(sales_amount_inr),2)          AS revenue_inr,
    ROUND(SUM(gross_profit_inr),2)          AS gross_profit_inr,
    ROUND(AVG(profit_margin_pct),2)         AS avg_margin_pct,
    ROUND(AVG(sales_amount_inr),2)          AS avg_order_value,
    ROUND(SUM(CASE WHEN return_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS return_rate_pct,
    ROUND(AVG(days_to_deliver),2)           AS avg_delivery_days
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY YEAR(order_date)
ORDER BY YEAR(order_date);

-- 12.6 Payment method share by year
SELECT
    YEAR(order_date)                        AS [year],
    payment_method,
    COUNT(*)                                AS orders,
    ROUND(COUNT(*)*100.0/SUM(COUNT(*)) OVER (PARTITION BY YEAR(order_date)),2) AS yearly_share_pct
FROM orders
WHERE order_status != 'Cancelled'
GROUP BY YEAR(order_date), payment_method
ORDER BY YEAR(order_date), COUNT(*) DESC;

-- ============================================================
-- END OF PROJECT 04: E-COMMERCE SALES ANALYTICS
-- ============================================================
