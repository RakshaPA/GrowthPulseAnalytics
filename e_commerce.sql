-- ═══════════════════════════════════════════════════════════
--  AUTOMATED GROWTH AUDIT — SQL AGGREGATION QUERIES
--  Compatible: PostgreSQL / BigQuery / Snowflake / MySQL
-- ═══════════════════════════════════════════════════════════

-- ─────────────────────────────────────────────────────────────
-- 1. MONTHLY KPI AGGREGATION
-- ─────────────────────────────────────────────────────────────
SELECT
  DATE_TRUNC('month', order_date)                              AS month,
  SUM(net_revenue)                                             AS total_revenue,
  COUNT(DISTINCT order_id)                                     AS total_orders,
  COUNT(DISTINCT customer_id)                                  AS unique_customers,
  ROUND(SUM(net_revenue)/NULLIF(COUNT(DISTINCT order_id),0),2) AS avg_order_value,
  SUM(marketing_spend)                                         AS total_mkt_spend,
  ROUND(SUM(net_revenue)/NULLIF(SUM(marketing_spend),0),2)    AS roas,
  SUM(is_returned)                                             AS total_returns
FROM ecommerce_orders
GROUP BY 1 ORDER BY 1;

-- ─────────────────────────────────────────────────────────────
-- 2. CUSTOMER ACQUISITION COST (CAC)
-- ─────────────────────────────────────────────────────────────
SELECT
  DATE_TRUNC('month', order_date)                              AS month,
  SUM(marketing_spend)                                         AS mkt_spend,
  COUNT(DISTINCT CASE WHEN is_new_customer=1
        THEN customer_id END)                                  AS new_customers,
  ROUND(SUM(marketing_spend)/
    NULLIF(COUNT(DISTINCT CASE WHEN is_new_customer=1
           THEN customer_id END),0),2)                         AS cac
FROM ecommerce_orders
GROUP BY 1 ORDER BY 1;

-- ─────────────────────────────────────────────────────────────
-- 3. CONVERSION RATE BY MONTH
-- ─────────────────────────────────────────────────────────────
SELECT
  DATE_TRUNC('month', order_date)                              AS month,
  SUM(website_visits)                                          AS total_visits,
  COUNT(DISTINCT order_id)                                     AS total_orders,
  ROUND(COUNT(DISTINCT order_id)::FLOAT/
        NULLIF(SUM(website_visits),0)*100,2)                   AS conversion_rate_pct
FROM ecommerce_orders
GROUP BY 1 ORDER BY 1;

-- ─────────────────────────────────────────────────────────────
-- 4. CATEGORY: GROSS MARGIN, RETURN RATE, ROAS
-- ─────────────────────────────────────────────────────────────
SELECT
  category,
  ROUND(SUM(net_revenue),2)                                    AS total_revenue,
  ROUND((SUM(net_revenue)-SUM(cogs))/
        NULLIF(SUM(net_revenue),0)*100,1)                      AS gross_margin_pct,
  COUNT(DISTINCT order_id)                                     AS orders,
  ROUND(SUM(net_revenue)/NULLIF(COUNT(DISTINCT order_id),0),2) AS aov,
  ROUND(SUM(is_returned)::FLOAT/COUNT(DISTINCT order_id)*100,1)AS return_rate_pct,
  ROUND(SUM(net_revenue)/NULLIF(SUM(marketing_spend),0),2)    AS roas
FROM ecommerce_orders
GROUP BY 1 ORDER BY total_revenue DESC;

-- ─────────────────────────────────────────────────────────────
-- 5. REPEAT PURCHASE RATE
-- ─────────────────────────────────────────────────────────────
SELECT
  DATE_TRUNC('month', order_date)                              AS month,
  COUNT(DISTINCT customer_id)                                  AS total_customers,
  COUNT(DISTINCT CASE WHEN is_new_customer=0
        THEN customer_id END)                                  AS repeat_customers,
  ROUND(COUNT(DISTINCT CASE WHEN is_new_customer=0
        THEN customer_id END)::FLOAT/
        NULLIF(COUNT(DISTINCT customer_id),0)*100,1)           AS repeat_rate_pct
FROM ecommerce_orders
GROUP BY 1 ORDER BY 1;

-- ─────────────────────────────────────────────────────────────
-- 6. INVENTORY TURNOVER
-- ─────────────────────────────────────────────────────────────
SELECT
  category,
  ROUND(AVG(inventory_units),0)                                AS avg_inventory,
  ROUND(SUM(cogs),2)                                           AS total_cogs,
  ROUND(SUM(cogs)/NULLIF(AVG(inventory_units),0),2)            AS inventory_turnover
FROM ecommerce_orders
GROUP BY 1 ORDER BY inventory_turnover DESC;

-- ─────────────────────────────────────────────────────────────
-- 7. CAMPAIGN ROI RANKING
-- ─────────────────────────────────────────────────────────────
SELECT
  campaign,
  ROUND(SUM(net_revenue),2)                                   AS total_revenue,
  ROUND(SUM(marketing_spend),2)                               AS total_spend,
  ROUND(SUM(net_revenue)/NULLIF(SUM(marketing_spend),0),2)   AS roas,
  ROUND(SUM(marketing_spend)/NULLIF(
        COUNT(DISTINCT CASE WHEN is_new_customer=1
        THEN customer_id END),0),2)                           AS cac
FROM ecommerce_orders
GROUP BY 1 ORDER BY roas DESC;

-- ─────────────────────────────────────────────────────────────
-- 8. ANOMALY DETECTION — Revenue Drop > 5% MoM
-- ─────────────────────────────────────────────────────────────
WITH monthly AS (
  SELECT DATE_TRUNC('month',order_date) AS month,
         SUM(net_revenue) AS revenue, SUM(marketing_spend) AS spend
  FROM ecommerce_orders GROUP BY 1
),
lagged AS (
  SELECT month, revenue, spend,
         LAG(revenue) OVER (ORDER BY month) AS prev_revenue
  FROM monthly
)
SELECT month,
  ROUND((revenue-prev_revenue)/NULLIF(prev_revenue,0)*100,1) AS rev_pct_change,
  ROUND(revenue/NULLIF(spend,0),2)                           AS roas,
  CASE
    WHEN (revenue-prev_revenue)/NULLIF(prev_revenue,0)<-0.05
    THEN 'CRITICAL: Revenue dropped >5% MoM'
    WHEN revenue/NULLIF(spend,0)<6
    THEN 'WARNING: ROAS below 6x threshold'
    ELSE 'OK'
  END AS anomaly_flag
FROM lagged
WHERE prev_revenue IS NOT NULL ORDER BY 1;

-- ─────────────────────────────────────────────────────────────
-- 9. CUSTOMER LIFETIME VALUE (CLV) PROXY
-- ─────────────────────────────────────────────────────────────
SELECT
  customer_id,
  COUNT(DISTINCT order_id)                AS total_orders,
  ROUND(SUM(net_revenue),2)               AS lifetime_revenue,
  ROUND(AVG(net_revenue),2)               AS avg_order_value,
  MIN(order_date)                         AS first_order,
  MAX(order_date)                         AS last_order,
  CASE WHEN COUNT(DISTINCT order_id)>=3  THEN 'High Value'
       WHEN COUNT(DISTINCT order_id)=2   THEN 'Repeat'
       ELSE 'One-Time' END                AS customer_segment
FROM ecommerce_orders
GROUP BY 1 ORDER BY lifetime_revenue DESC LIMIT 100;

-- ─────────────────────────────────────────────────────────────
-- 10. H1 vs H2 ROOT CAUSE COMPARISON
-- ─────────────────────────────────────────────────────────────
SELECT
  CASE WHEN EXTRACT(MONTH FROM order_date)<=6
       THEN 'H1-2024' ELSE 'H2-2024' END                     AS half_year,
  ROUND(SUM(net_revenue),2)                                   AS revenue,
  ROUND(SUM(net_revenue)/NULLIF(SUM(marketing_spend),0),2)   AS roas,
  ROUND(COUNT(DISTINCT order_id)::FLOAT/
        NULLIF(SUM(website_visits),0)*100,2)                  AS conv_rate_pct,
  SUM(is_returned)                                            AS returns,
  ROUND((SUM(net_revenue)-SUM(cogs))/
        NULLIF(SUM(net_revenue),0)*100,1)                     AS gross_margin_pct
FROM ecommerce_orders
GROUP BY 1 ORDER BY 1;