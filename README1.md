# Introduction
## 1. Context

This project focuses on analyzing transaction and payment data for an e-wallet company. The analysis is divided into two main parts: Exploratory Data Analysis (EDA) and Data Wrangling. The datasets used include payment_report.csv, product.csv, and transactions.csv, each providing essential insights into the company's performance, product offerings, and transaction behaviors.

## 2. Dataset:
product.csv (product information)
|Variable|Description|
|-|-|
|product_id|Unique identifier for each product.|
|category|Category code or name representing the type of product.|
|team_own|Team responsible for the product.|

payment_report.csv (monthly payment volume of products)
|Variable|Description|
|-|-|
|report_month|Month and year for the payment report (e.g., "2023-01" for January 2023).|
|payment_group|Group type of the payment, such as "payment" or "refund."|
|product_id|ID of the product involved in the payment transaction.|
|source_id|ID representing the source of the transaction.|
|volume|Total transaction volume for the given product and source in the specified month.|  

transaction.csv (transactions information)
|Variable|Description|
|-|-|
|transaction_id|Unique identifier for each transaction.|
|merchant_id|ID of the merchant involved in the transaction.|
|volume|The monetary amount of the transaction.|
|transType|Code representing the type of transaction (e.g., payment, withdrawal).|
|transStatus|Status code indicating the success or failure of the transaction.|
|sender_id|ID of the sender initiating the transaction.|
|receiver_id|ID of the receiver for the transaction.|
|extra_info|Additional information related to the transaction, if applicable.|
|timeStamp|Timestamp indicating the time when the transaction occurred.|

# Questions
## Part 1: EDA
Do EDA Task:
- Df payment_enriched (Merge payment_report.csv and product.csv).
- Df transactions.

Suggestions:
- Missing values: Check missing values on rows, suggest next step, actions
- Duplicates: Check duplicates records on rows, suggest next step, actions
- Incorrect data type: Identify on each column with inccorect datatype, suggest next step, actions
- Inccorrect values: Identify on each column with incorrect values, suggest next stepm actions
## Part 2: Data Wrangling
Using payment_report.csv and product.csv to:
- Top 3 Product_id with the highest volume.
- Given that 1 product_id is only owed by 1 team, are there any abnormal products against this rule.
- Find the team has had the lowest performance (lowest volume) since Q2.2023. Find the category that contributes the least to that team.
- Find the contribution of source_ids of refund transactions (payment_group = ‘refund’), what is the source_id with the highest contribution?

Using transactions.csv
- Classify each transactions in this file based on several criteria (Bank Transfer Transaction, Withdraw Money Transaction, Top Up Money Transaction,..)
- From each transaction type (excluding invalid transactions): find the number of transactions, volume, senders and receivers.

