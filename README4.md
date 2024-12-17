# I. Introduction
## 1. Context
This application is designed to help businesses (telecom industry) predict customer churn, which refers to the likelihood of customers discontinuing their services or subscriptions. By calculating churn rate ,identifying customers at risk of churn with the various churn reasons, businesses can implement targeted retention strategies to improve customer satisfaction and reduce revenue loss.
## 2. Dataset
Dataset relate to user churn in telecom field, containing a dimension table with information on users of a telecommunications network company. The table contains fields for demographic informations, revenue as well as churned user.
## 3. Data Dictionary
|Field|Meaining|
|-|-|
|Customer ID|ID of Customers.|
|Churn Label|Is It Churned Customer?.|
|Account Length (in months)|Duration since the customer opened an account until now (in months).|
|Local Calls|Numbers of domestic calls.|
|Local Mins|Domestic calls' duration.|
|Intl Calls|Numbers of International calls.|
|Intl Mins|International calls' duration.|
|Intl Active|Whether the customer is subscribed to international calls.|
|Extra International Charges|Additional charges for international calls.|
|Customer Service Calls|Number of customer service calls made.|
|Avg Monthly GB Download|Average monthly GB downloaded.|
|Unlimited Data Plan|Whether the customer uses an unlimited data plan.|
|Extra Data Charges| Additional data charges.|
|State|Customer’s state.|
|Phone Number|Customer's phone number.|
|Gender|Customer's gender.|
|Age|Customer's age.|
|Under 30|Whether the customer is under 30 years old.|
|Group|Whether the customer is subscribed to group services.|
|Number of Customers in Group|Number of people in a group.|
|Device Protection & Online Backup|Whether the customer uses data security and online backup services.|
|Contract Type|Type of contract.|
|Payment Method|Payment method.|
|Monthly Charge|Monthly fee paid by the customer.|
|Total Charges|Total amount paid by the customer.|
|Churn Category|Churn category.|
|Churn Reason|Reason for churn.|
## 4. Requirement
Create overview dashboard for Managers to have a overall picture about churn situation of users, as well as churn reasons in each churn group, thereby suggesting a approrpriate solutions to improve this problems
# II. Visualization by Power BI

![image](https://github.com/user-attachments/assets/b85fcea6-a2cf-43de-9c47-1977a7d11d72)

![image](https://github.com/user-attachments/assets/074581ad-23eb-45f6-9a52-e3aff0784347)

![image](https://github.com/user-attachments/assets/99756856-982f-4ffe-a0aa-2b29f24429ec)

# III. Overall and Insights
### 1. Customer Overview:
As can be seen from dashboard that
- Churned customer constitued 1,769 customers compared with 6,687 of total, with the churn rate is nearly 27% which indicates a significant problem requiring possible solutions => the amount of both genders customers still be balanced.

- Gender critiria: show the same proportion in both gender for both total and churned customers, while age group is distributed clearly between youngsters and the elder customers, within the over-30 aged group show the higher proportion than the aonther in both group of customers (over 80%) => ALthough the olders can be the most active customers, we must pay attention on this group to find out reason leading to the high churn rate.

- Contract type: The monthly packages dominates others in the number of customers (51%), while the rate of monthly packages of churned users witnessed the highest proportion (87.9) => Have some issues in this packages that need to be focused on.

### 2. Churn reason
Top 3 churn reasons: the first reason that constitute highest amount is Competitor made better offer (303), followed by Competitor had better devices and Attitude of support person, are 297 and 203 respectively => Need to promotion campaigns, improve and innovate devices as well as propose training for staff in supporting or caring customers.
### 3. Recommendation.
From these insights analysed above:
- Pay attention on older group to have a clear perspective about their churn reason, caring about post-service to ensure customer satisfaction
- Propose several appropriate solutions to mitigate issue about packages type
- Knowing and researching competitor, propose another promotion or offer campaign to encourage customer, collaborate with R&D team to innovate products, and training sales team to support customers carefully

