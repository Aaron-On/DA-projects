--Q1: calculate total visit, pageview, transaction for Jan, Feb and March 2017 (order by month)
SELECT
  format_date('%Y%m', (parse_date('%Y%m%d',date))) as date_time,
  SUM(totals.visits) as visits,
  SUM(totals.pageviews) as pageviews,
  SUM(totals.transactions) as transactions
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_2017*`
WHERE _table_suffix between '0101' and '0331'
GROUP BY format_date('%Y%m', (parse_date('%Y%m%d',date)))
ORDER BY date_time;

--Q2: Bounce rate per traffic source in July 2017 (Bounce_rate = num_bounce/total_visit) (order by total_visit DESC)
SELECT  
  trafficSource.source,
  SUM(totals.visits) as total_visits,
  SUM(totals.bounces) as total_no_of_bounces,
  ROUND((SUM(totals.bounces)/SUM(totals.visits))*100.0, 3) as bounce_rate
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_2017*` 
WHERE _table_suffix between '0701' and '0731'
GROUP BY trafficSource.source
ORDER BY total_visits DESC;

--Q3: Revenue by traffic source by week, by month in June 2017
WITH source_by_month as
(SELECT  
  'Month' as time_type,
  format_date('%Y%m', (parse_date('%Y%m%d',date))) as time,
  trafficSource.source as source,
  ROUND(SUM(product.productRevenue)/1000000, 4) as revenue
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_2017*`,
UNNEST(hits) as hits,
UNNEST(hits.product) as product
WHERE _table_suffix between '0601' and '0630'
AND product.productRevenue IS NOT NULL
GROUP BY time, trafficSource.source),

source_by_week as
(SELECT  
  'Month' as time_type,
  format_date('%Y%W', (parse_date('%Y%m%d',date))) as time,
  trafficSource.source as source,
  ROUND(SUM(product.productRevenue)/1000000, 4) as revenue
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_2017*`,
UNNEST(hits) as hits,
UNNEST(hits.product) as product
WHERE _table_suffix between '0601' and '0630'
AND product.productRevenue IS NOT NULL
GROUP BY time, trafficSource.source)

SELECT *
FROM source_by_month
UNION ALL
SELECT *
FROM source_by_week
ORDER BY revenue DESC;

--Q4: Average number of pageviews by purchaser type (purchasers vs non-purchasers) in June, July 2017.
WITH num_user as 
(SELECT  
  distinct format_date('%Y%m', (parse_date('%Y%m%d',date))) as month,
  SUM(CASE 
        WHEN totals.transactions >= 1 and product.productRevenue IS NOT NULL THEN totals.pageviews END) as total_pageviews_purchase,
  SUM(CASE
        WHEN totals.transactions IS NULL and product.productRevenue IS NULL THEN totals.pageviews END) as total_pageviews_non_purchase,
  COUNT(distinct CASE 
                    WHEN totals.transactions >= 1 and product.productRevenue IS NOT NULL THEN fullVisitorId END) as total_num_userpurchase,
  COUNT(distinct CASE 
                    WHEN totals.transactions IS NULL and product.productRevenue IS NULL THEN fullVisitorId END) as total_num_usernonpur
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_2017*`,
UNNEST (hits) as hits,
UNNEST (hits.product) as product
WHERE _table_suffix between '0601' and '0731' 
GROUP BY month)
SELECT
  month,
  ROUND(total_pageviews_purchase/total_num_userpurchase,7) as avg_pageviews_purchase,
  ROUND(total_pageviews_non_purchase/total_num_usernonpur,7) as avg_pageviews_non_purchase
FROM num_user
ORDER BY month;

--Q5: Average number of transactions per user that made a purchase in July 2017.
SELECT 
  format_date('%Y%m', (parse_date('%Y%m%d',date))) as Month,
  ROUND(SUM(totals.transactions)/COUNT(distinct fullVisitorID),9) as Avg_total_transactions_per_user
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_201707*`,
UNNEST (hits) as hits,
UNNEST (hits.product) as product
WHERE totals.transactions >=1
AND product.productRevenue IS NOT NULL
GROUP BY  format_date('%Y%m', (parse_date('%Y%m%d',date)));

--Q6: Average amount of money spent per session. Only include purchaser data in July 2017.
SELECT 
  format_date('%Y%m', (parse_date('%Y%m%d',date))) as Month,
  ROUND((SUM(product.productRevenue)/ SUM(totals.visits))/1000000,3) as avg_revenue_by_user_per_visit
 FROM `bigquery-public-data.google_analytics_sample.ga_sessions_201707*`,
 UNNEST (hits) as hits,
 UNNEST (hits.product) as product 
 WHERE totals.transactions IS NOT NULL
 AND product.productRevenue IS NOT NULL
 GROUP BY Month;

 --Q7: Other products purchased by customers who purchased product "YouTube Men's Vintage Henley" in July 2017. Output should show product name and the quantity was ordered.
WITH cust_purchase_YTMVhenley as 
(SELECT
  distinct fullVisitorId,
  product.v2ProductName
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_201707*`,
UNNEST (hits) as hits,
UNNEST (hits.product) as product
WHERE product.productRevenue IS NOT NULL
AND product.v2productName = "YouTube Men's Vintage Henley")

SELECT
  product.v2productName another_purchased_products,
  SUM(product.productQuantity) as quantity
FROM cust_purchase_YTMVhenley as cust_ID
LEFT JOIN `bigquery-public-data.google_analytics_sample.ga_sessions_201707*` as full_prod
ON cust_ID.fullVisitorId = full_prod.fullVisitorId,
UNNEST (hits) as hits,
UNNEST (hits.product) as product
WHERE product.productRevenue IS NOT NULL
AND product.v2productName <> "YouTube Men's Vintage Henley" 
GROUP BY product.v2productName
ORDER BY quantity DESC;

--Q8:Calculate cohort map from product view to addtocart to purchase in Jan, Feb and March 2017. For example, 100% product view then 40% add_to_cart and 10% purchase. Add_to_cart_rate = number product add to cart/number product view. Purchase_rate = number product purchase/number product view. The output should be calculated in product level.
WITH addtocart_purchase as
(SELECT  
  format_date('%Y%m', (parse_date('%Y%m%d',date))) as Month,
  COUNT(CASE 
          WHEN hits.eCommerceAction.action_type = '2' AND product.isImpression IS NULL THEN fullVisitorId END) as num_product_view,
  COUNT(CASE 
          WHEN hits.eCommerceAction.action_type = '3' AND product.isImpression IS NULL THEN fullVisitorId END) as num_addtocart,
  COUNT(CASE 
          WHEN hits.eCommerceAction.action_type = '6' AND product.isImpression IS NULL AND product.productRevenue IS NOT NULL THEN fullVisitorId END) as num_purchase
FROM `bigquery-public-data.google_analytics_sample.ga_sessions_2017*`,
UNNEST (hits) as hits,
UNNEST (hits.product) as product
WHERE _table_suffix between '0101' and '0331'
GROUP BY Month)
SELECT
  Month,
  num_product_view,
  num_addtocart,
  num_purchase,
  ROUND((num_addtocart/num_product_view)*100.0,2) as add_to_cart_rate,
  ROUND((num_purchase/num_product_view)*100.0,2) as purchase_rate
FROM  addtocart_purchase
ORDER BY Month;
