# I. Introduction
The AdventureWorks2019 SQL Project leverages the AdventureWorks2019 dataset, a comprehensive business database designed to simulate the operations of a fictional global manufacturing company. Exploring this dataset by using SQL on GoogleBigQuery, find the solutions for following questions realating to AdventureWorks2019 Project
## 1. Context
The AdventureWorks2019 dataset stimulate real-world business operations.It represents a fictitious global manufacturing company, including data related to sales, production, products,..etc for analysis and get insights.
## 2. Dataset
- Sales Data: SalesOrderHeader, SalesOrderDetail, Sales.SpecialOffer
- Production Data: Production.Product, ProductionSubcategory, Production.WorkOrder
- Purchase Data: Purchasing.PurchaseOrderHeader
# II. Exploring Dataset
### Query 1: Calc Quantity of items, Sales value & Order quantity by each Subcategory in L12M

SQL code:

![image](https://github.com/user-attachments/assets/f89b6a12-024c-44c7-a39f-4ca080a83247)

Query Results:

![image](https://github.com/user-attachments/assets/79eef940-2de1-4dd1-ab32-683798357021)

### Query 2: Calc % YoY growth rate by SubCategory & release top 3 cat with highest grow rate. Can use metric: quantity_item. Round results to 2 decimal

SQL code:

![image](https://github.com/user-attachments/assets/55195eae-c021-4de6-9245-34b5a7c62f34)
![image](https://github.com/user-attachments/assets/3ec1ac06-1d61-408e-aea1-f9aaa33c9633)

Query Results:

![image](https://github.com/user-attachments/assets/08c66ae2-f0b5-4332-951c-6a1fdf5b024e)

### Query 3: Ranking Top 3 TeritoryID with biggest Order quantity of every year. If there's TerritoryID with same quantity in a year, do not skip the rank number

SQL code:

![image](https://github.com/user-attachments/assets/c58505cc-c0d6-480f-a3e0-a3d24df1605e)

Query Results:

![image](https://github.com/user-attachments/assets/29613979-e0ac-4f66-b55d-4c6f2586bf20)

### Query 4: Calc Total Discount Cost belongs to Seasonal Discount for each SubCategory

SQL code:

![image](https://github.com/user-attachments/assets/799ace78-2f73-44e1-971f-c5f7f0126a91)

Query Results:

![image](https://github.com/user-attachments/assets/c4767967-4e92-45ac-94f8-a53041f6ec3f)

### Query 5: Retention rate of Customer in 2014 with status of Successfully Shipped (Cohort Analysis).

SQL code:

![image](https://github.com/user-attachments/assets/ca48497d-10e9-4cd7-a6c7-4bac6e4c494f)

Query Results:

![image](https://github.com/user-attachments/assets/ce73de61-e6dc-450e-a865-58e153b725df)

### Query 6: Trend of Stock level & MoM diff % by all product in 2011. If %gr rate is null then 0. Round to 1 decimal.

SQL code:

![image](https://github.com/user-attachments/assets/f360244b-d480-43c0-b05f-7803347b4caf)

Query Results:

![image](https://github.com/user-attachments/assets/1c3e1c05-f04c-4b8c-8e8a-e7be09cb5ca5)

### Query 7: Calc Ratio of Stock / Sales in 2011 by product name, by month. Order results by month desc, ratio desc. Round Ratio to 1 decimal mom yoy.

SQL code:

![image](https://github.com/user-attachments/assets/7f48d973-c6c8-4067-b79d-61743b6bd7e3)

Query Results:

![image](https://github.com/user-attachments/assets/94c8e5c1-f60b-4913-acf0-2487db9499a6)

### Query 8: No of order and value at Pending status in 2014.

SQL code:

![image](https://github.com/user-attachments/assets/a3a6003b-bfd5-4160-8ae0-fdbd6b073947)

Query Results:

![image](https://github.com/user-attachments/assets/8245e3b6-31d7-421e-a015-d0e0f3f9e2b6)

# III. Conclusion
The Exploration of AdvantureWorks2019 project by using Google BigQuery enhance foundation knowledge and gain valuable information about products, Orders, Quantity of products,... which help company get insights and propose appropriate campaigns in the future. Besides, The scalability and performance of BigQuery allow for processing large datasets efficiently, enabling organizations to make data-driven decisions 






















