
# I. Introduction
The AdventureWorks Power BI Project is designed to analyze and visualize the business performance of the fictional AdventureWorks company, a global manufacturer and retailer of bicycles and cycling accessories. The project leverages the company's transactional and operational data to provide actionable insights for stakeholders.
## 1.Context
AdventureWorks database supports standard online transaction processing scenarios for a fictitious bicycle manufacturer - Adventure Works Cycles. 
Within, Production Department is responsible for receiving manufacturing's request from the Company's Planning Department. After Procurement placing orders and receiving raw materials into the warehouse, the production process begins
According to the plan, after production have been completed, the products will be stored at different locations to facilitate storage and distribute to customers. However, the actual production completed time may not meet the plan. During the inspection process for warehousing, there is a possibility of damaged goods that will be removed to ensure that only high-quality products are stored for sale
## 2. Dataset
The AdventureWorks dataset is a relational database representing a global manufacturing company. It contains tables capturing  product tables , ideal for creating interactive Power BI dashboards and reports.
## 3. Data Dictionary
The following tables that is imported from Adventure project SQL are:

- Product_Inventory
- Product_Product
- ProductCategory
- Production_BillofMaterials
- Production_Location
- Production_ScrapReason
- Production_WorkOrder
- Production_WorkOrderDetail

## 4. Requirement
Create an operation dashboard providing insights for Production Director overall picture about company production's performance to develop decision and operate factory effectively. Besides, give some more recommedations.
# II. Visualization by Power BI
![image](https://github.com/user-attachments/assets/b38440f1-8246-4194-8c8d-5740bdf01b93)

![image](https://github.com/user-attachments/assets/9a032a09-e2b1-44a3-9fcf-91da3e3307cc)

![image](https://github.com/user-attachments/assets/dbfe8858-8939-4da6-b944-1ce943c347bc)

# III. Overall and Insights
### 1.Insight

- The number of product accompanied remarkably by bike and components , while the number of the latter is significantly higher than the another in production  => Two out of four products (bikes and component) which is in high demand manufactured in-house in order to save cost , within components  is mainly produced for assembling bikes and another one.
  
- Lead-time couldn't meet target during week and month period => have some bottlenecks in certain production stages during process
Production is concentrated mainly between February  and October ( in the middle of the year) and then witnessed the significant downward trend.=> Fiscal year could be start at the October => Focus on high-demands products, could meet totally customer's demand, have a healthy production activity (by many reasons such as rising orders, rising customer's demand) get diverse markets, be less dependable on single product line.

- Regarding to Production Analysis: Total on-time product is probably high ( 3533K vs 4508K in total) , Yield rate is 99.76% , while %On-Time Order is just 68.76% <80% => Production still be meet on-time products that can balance on different kind of categories , with the high of yield rate number indicating consistent quality  in production. However the on-time order is fairly good , that can be due to inefficiencies in material usage, quality control gaps and control works order for each kind of products.
  
- Besides, the number of products and number of orders fluctuates steadily in over medium level over month => be good production and work orders
Moreover, Wheels and Derailleurs are two popular sub-category product in-house manufacturing, within BB Ball Bearing is the most efficiency product in process.

- The production cost of bikes still be remained unchanged during month time over many year, while the another of components fluctuates significantly => The bike keep on efficiency in material usage, but the component get on bottlenecks in material and find another supplier for replacing it.
  
- Regarding to Scrap Analysis, Scrap rate just hold 0.236% and scrap order is 1.004% with Paint process failed is the main reason during production process => quality issue keep good-control, have no any issues such as machinery , lack of operation training.

### 2.Recommendation

- Depend on each kind of customer demand in order to produce suitable products.
  
- Track each production stages associated with products to seek for common steps where delays happens usually, explore solutions like recheck material , stock, upgrading equipments relocating resource.
  
- Keep going on growth trend during specific period in year, prioritize high-demand focus, finding more various markets.
  
- Prioritize on high-demand products, consider on quality assurance , inspection process, as well as machinery maintenance and skill operation development in order to keep high yield rate. Moreover, balance on work orders, time management, distribute  schedule effectively in order to meet the workflow, prevent increase late-time orders.
  
- Identify the common reason during production process, stricter quality check, create training session to develop skill operation, workshops and meeting is needed for show off the negative impact in failure production.
  
- Keep going on low-level scrap products, focus on quality check to enhance the stability of products, prevent focusing on one kind of products.

