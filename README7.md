# I.Introduction
eCommercee dataset Exploration project using SQL in Google BigQuery Platform. This data set is basef on bigquery-public-data.google_analytics_sample.ga_sessions_2017. This project aim to practice and apply lessons on BigQuery Platform to enhance foundation of SQL.
## 1. Context
In the digital landscape, businesses rely heavily on website analytics to understand how users interact with their platforms. Google Analytics captures vital information about user sessions, including:

- Pageviews and user interactions.
- E-commerce transactions, including revenue, products, ecommerce actions, and product quantity,..
## 2.Dataset
For each Analytics view that is enabled for BigQuery integration, a dataset is added using the view ID as the name.
## 3.Data Dictionary
|Field Name|Data Type|Description|
|-|-|-|
|fullVisitorId|STRING|The unique visitor ID.|
|date|STRING|The date of the session in YYYYMMDD format.|
|totals|RECORD|This section contains aggregate values across the session.|
|totals.bounces|INTEGER|Total bounces (for convenience). For a bounced session, the value is 1, otherwise it is null.|
|totals.hits|INTEGER|Total number of hits within the session.|
|totals.pageviews|INTEGER|Total number of pageviews within the session.|
|totals.visits|INTEGER|The number of sessions (for convenience). This value is 1 for sessions with interaction events. The value is null if there are no interaction events in the session.|
|totals.transactions|INTEGER|Total number of ecommerce transactions within the session.|
|trafficSource.source|STRING|The source of the traffic source. Could be the name of the search engine, the referring hostname, or a value of the utm_source URL parameter.|
|hits|RECORD|This row and nested fields are populated for any and all types of hits.|
|hits.eCommerceAction|RECORD|This section contains all of the ecommerce hits that occurred during the session. This is a repeated field and has an entry for each hit that was collected.|
|hits.eCommerceAction.action_type|STRING|The action type. Click through of product lists = 1, Product detail views = 2, Add product(s) to cart = 3, Remove product(s) from cart = 4, Check out = 5, Completed purchase = 6, Refund of purchase = 7, Checkout options = 8, Unknown = 0. Usually this action type applies to all the products in a hit, with the following exception: when hits.product.isImpression = TRUE, the corresponding product is a product impression that is seen while the product action is taking place (i.e., a ""product in list view"").|
|hits.product|RECORD|This row and nested fields will be populated for each hit that contains Enhanced Ecommerce PRODUCT data.|
|hits.product.productQuantity|INTEGER|The quantity of the product purchased.|
|hits.product.productRevenue|INTEGER|The revenue of the product, expressed as the value passed to Analytics multiplied by 10^6 (e.g., 2.40 would be given as 2400000).|
|hits.product.productSKU|STRING|Product SKU.|
|hits.product.v2ProductName|STRING|Product Name.|
# II.Exploring data
### Query 1: Calculate total visit, pageview, transaction for Jan, Feb and March 2017 (order by month)

SQL code:

![image](https://github.com/user-attachments/assets/62339886-280a-4123-a777-021f8e9a5e66)

Query Results:

![image](https://github.com/user-attachments/assets/7d3ca160-6c41-4ba5-9b7c-9f5d0ec08a50)

### Query 2: Bounce rate per traffic source in July 2017 (Bounce_rate = num_bounce/total_visit) (order by total_visit DESC)

SQL code:

![image](https://github.com/user-attachments/assets/0a917998-c8cc-4a6a-a580-2337b93f95fa)

Query Results

![image](https://github.com/user-attachments/assets/c0b2cb91-c2b9-4a57-97ff-2b452f6b81f5)

### Query 3: Revenue by traffic source by week, by month in June 2017

SQl code:

![image](https://github.com/user-attachments/assets/5f275921-7074-4d55-b975-d0ac796b1579)

Query Results:

![image](https://github.com/user-attachments/assets/a8cc7175-7b2e-419c-a0ac-66aeba8a6f04)

### Query 4: Average number of pageviews by purchaser type (purchasers vs non-purchasers) in June, July 2017.

SQL code:

![image](https://github.com/user-attachments/assets/6cf6a9e0-1a8c-41ef-bcaf-3038f60ed8e9)

Query Results:

![image](https://github.com/user-attachments/assets/218d2a67-0284-411c-b43a-a5d958d06ecd)

### Query 5: Average number of transactions per user that made a purchase in July 2017.

SQL code:

![image](https://github.com/user-attachments/assets/a227a6c8-6d52-4c4e-be12-9703e7ddbfb0)

Query Results:

![image](https://github.com/user-attachments/assets/ebfb4b76-c626-407f-acbe-ff469171fb13)

### Query 6: Average amount of money spent per session. Only include purchaser data in July 2017.

SQL code:

![image](https://github.com/user-attachments/assets/bb07798a-a912-440e-a724-8aead03c05ea)

Query Results:

![image](https://github.com/user-attachments/assets/458bb0e3-cfb6-4c28-a944-216d91c9e530)


### Query 7: Other products purchased by customers who purchased product "YouTube Men's Vintage Henley" in July 2017. Output should show product name and the quantity was ordered.

SQL code:

![image](https://github.com/user-attachments/assets/bd5b66fb-2ba3-4d89-be4d-de00aba73e7e)

Query Results:

![image](https://github.com/user-attachments/assets/0cb1e9e6-85f4-4ffe-a999-9b4d34e6a040)

### Query 8: Calculate cohort map from product view to addtocart to purchase in Jan, Feb and March 2017. For example, 100% product view then 40% add_to_cart and 10% purchase. Add_to_cart_rate = number product  add to cart/number product view. Purchase_rate = number product purchase/number product view. The output should be calculated in product level.

SQL code:

![image](https://github.com/user-attachments/assets/8eb1fc76-f19b-4762-bade-f987f4a412b4)

Query Results:

![image](https://github.com/user-attachments/assets/a75988c9-db4c-4516-9f1e-49e79b592772)

# III. Conclusion
Knowing how to query on BigQuery platform, utilize sql in ecommerce project in order to have insights about company's performance.
Optimize tasks to propose appropriate strategies in the future for business


  





















