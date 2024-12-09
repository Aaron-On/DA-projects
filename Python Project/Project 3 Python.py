import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import datetime as dt
import warnings
warnings.filterwarnings('ignore')
#import excel file
ecommerce_retail = pd.read_excel(
    r'C:\\Users\\Admin\\Desktop\\DA books\\DA lesson\\Python\\Final_project_RFM\\ecommerce retail.xlsx'
)
Segmentation = pd.read_excel(
    r'C:\\Users\\Admin\\Desktop\\DA books\\DA lesson\\Python\\Final_project_RFM\\ecommerce retail.xlsx', sheet_name = 'Segmentation'
)
#check missing values, data type
print(ecommerce_retail.info())
print(ecommerce_retail[ecommerce_retail['Quantity'] < 0])
print(ecommerce_retail[ecommerce_retail['CustomerID'].isna()])



#Count number of missing values
print(ecommerce_retail.isna().sum())

ecommerce_retail['InvoiceNo'] = ecommerce_retail['InvoiceNo'].astype('string')
print(ecommerce_retail[ecommerce_retail['InvoiceNo'].str.contains("^C")])
print(ecommerce_retail[(ecommerce_retail['Quantity'] < 0) & (ecommerce_retail['InvoiceNo'].str.contains("^C"))])

#Drop unmeaning null values
ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[(ecommerce_retail['Description'].isna()) 
                     & (ecommerce_retail['CustomerID'].isna()) 
                     & (ecommerce_retail['UnitPrice']== 0)].index, axis = 0)

print(
    ecommerce_retail[
        (ecommerce_retail['Description'].isna()) 
        & (ecommerce_retail['CustomerID'].isna()) 
        & (ecommerce_retail['UnitPrice']== 0.0)
        ]
)

#change data type
ecommerce_retail['Description'] = ecommerce_retail['Description'].astype('string')
ecommerce_retail['Country'] = ecommerce_retail['Country'].astype('string')
ecommerce_retail['CustomerID'] = ecommerce_retail['CustomerID'].fillna(0).astype(int) 

print(ecommerce_retail.info())

#check incorrect values, outliers
print(ecommerce_retail.describe())
print(ecommerce_retail[ecommerce_retail['UnitPrice'] < 0])
print(ecommerce_retail[ecommerce_retail['Quantity'] < 0])
ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[ecommerce_retail['UnitPrice']< 0].index, axis = 0
)

ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[ecommerce_retail['Quantity']< 0].index, axis = 0
)

ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[ecommerce_retail['CustomerID']== 0].index, axis = 0
)

#Count number of duplicates
print(ecommerce_retail.duplicated().sum())

ecommerce_retail_removedup = ecommerce_retail.drop_duplicates()

print(ecommerce_retail_removedup.duplicated().sum())
print(ecommerce_retail_removedup)
print(ecommerce_retail_removedup.info())
print(ecommerce_retail_removedup[ecommerce_retail_removedup['InvoiceNo'].str.contains("^C")])
print(ecommerce_retail[(ecommerce_retail['Quantity'] < 0) & (ecommerce_retail['InvoiceNo'].str.contains("^C"))])

#Setup date for Recency at 31/12/2011
date_R = dt.datetime(2011,12,31)

#Calculate total revenue for each product
ecommerce_retail_removedup['Total_revenue'] = ecommerce_retail_removedup['Quantity'] * ecommerce_retail_removedup['UnitPrice']

#Calculate RFM model
RFM_model = ecommerce_retail_removedup.groupby('CustomerID').agg(
    Recency = ('InvoiceDate', lambda x: (date_R - x.max()).days),
    Frequency = ('InvoiceNo', 'count'),
    Monetary = ('Total_revenue', 'sum')
).reset_index()

print(RFM_model)

#Calculate rank for each customers based on quintile
RFM_model['Recency_Score'] = pd.qcut(RFM_model['Recency'], q = 5, labels=[5,4,3,2,1])
RFM_model['Frequency_Score'] = pd.qcut(RFM_model['Frequency'], q = 5, labels=[1,2,3,4,5])
RFM_model['Monetary_Score'] = pd.qcut(RFM_model['Monetary'], q = 5, labels=[1,2,3,4,5])

RFM_model['RFM_Score'] = RFM_model['Recency_Score'].astype(str) + RFM_model['Frequency_Score'].astype(str) + RFM_model['Monetary_Score'].astype(str)
RFM_model['RFM_Score'] = RFM_model['RFM_Score'].astype(int)
print(RFM_model)

#import sheet Segmentation to filter customer
# import numpy as np
# import pandas as pd
# import seaborn as sns
# import matplotlib.pyplot as plt
# import datetime as dt
# import warnings
# warnings.filterwarnings('ignore')
# #import excel file
# Segmentation = pd.read_excel(
#     r'C:\\Users\\Admin\\Desktop\\DA books\\DA lesson\\Python\\Final_project_RFM\\ecommerce retail.xlsx', sheet_name = 'Segmentation'
# )
print(Segmentation)

#Split one row to rows for each Segment
Segmentation['RFM Score'] = Segmentation['RFM Score'].str.split(',')
Segmentation_sepe = Segmentation.explode('RFM Score').reset_index(drop = True)
Segmentation_sepe['RFM Score'] = Segmentation_sepe['RFM Score'].astype(int)

#Print Seperated Segment
print(Segmentation_sepe)

#Join to filter RFM Score for each Segementation
RFM_Seg = pd.merge(RFM_model, Segmentation_sepe, left_on= 'RFM_Score', right_on= 'RFM Score', how = 'left')
print(RFM_Seg)

sns.catplot(x = "Segment", y ="Recency", data = RFM_Seg, kind = "bar")
plt.xticks(rotation = 90)
plt.show()
sns.catplot(x = "Segment", y ="Frequency", data = RFM_Seg, kind = "bar")
plt.xticks(rotation = 90)
plt.show()
sns.catplot(x = "Segment", y ="Monetary", data = RFM_Seg, kind = "bar")
plt.xticks(rotation = 90)
plt.show()


import numpy as np
import pandas as pd
import seaborn as sns
import os
import matplotlib.pyplot as plt
import datetime as dt
import warnings
warnings.filterwarnings('ignore')
ecommerce = pd.read_excel('ecommerce retail.xlsx')
print(ecommerce)
--


import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import datetime as dt
import warnings
warnings.filterwarnings('ignore')
#import excel file
ecommerce_retail = pd.read_excel(
    r'C:\\Users\\Admin\\Desktop\\DA books\\DA lesson\\Python\\Final_project_RFM\\ecommerce retail.xlsx'
)
Segmentation = pd.read_excel(
    r'C:\\Users\\Admin\\Desktop\\DA books\\DA lesson\\Python\\Final_project_RFM\\ecommerce retail.xlsx', sheet_name = 'Segmentation'
)

#check missing values, data type
print(ecommerce_retail.info())

#Count number of missing values
print("Number of missing values in: ")

print(ecommerce_retail.isna().sum())

#Remove null valuese in Customer ID
ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[ecommerce_retail['CustomerID'].isna()].index, axis = 0
)

#Recheck missing values
print("Number of missing values in: ")

print(ecommerce_retail.isna().sum())


#change data type
ecommerce_retail['InvoiceNo'] = ecommerce_retail['InvoiceNo'].astype('string')
ecommerce_retail['StockCode'] = ecommerce_retail['StockCode'].astype('string')
ecommerce_retail['Description'] = ecommerce_retail['Description'].astype('string')
ecommerce_retail['Country'] = ecommerce_retail['Country'].astype('string')
ecommerce_retail['CustomerID'] = ecommerce_retail['CustomerID'].astype(int) 

#Recheck data type
print(ecommerce_retail.info())

#check incorrect values, abnormal data
pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)

print(ecommerce_retail.describe())

#check negative values
pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)

print(ecommerce_retail[ecommerce_retail['UnitPrice'] < 0])

print(ecommerce_retail[ecommerce_retail['Quantity'] < 0])

#Count negative values in UnitPrice and Quantity
print((ecommerce_retail['UnitPrice'] < 0).sum())

print((ecommerce_retail['Quantity'] < 0).sum())

#Drop abnormal data, null values in Customer ID
ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[ecommerce_retail['UnitPrice']< 0].index, axis = 0
)

ecommerce_retail= ecommerce_retail.drop(
    ecommerce_retail[ecommerce_retail['Quantity']< 0].index, axis = 0
)
# ecommerce_retail= ecommerce_retail.drop(ecommerce_retail[ecommerce_retail['CustomerID']== 0].index, axis = 0)

#Recheck negative values
print((ecommerce_retail['UnitPrice'] < 0).sum())

print((ecommerce_retail['Quantity'] < 0).sum())

#Check outliers in UnitPrice col
sns.boxplot(ecommerce_retail, x = 'UnitPrice')  

#Check outliers in Quantity col
sns.boxplot(ecommerce_retail, x = 'Quantity') 

#Calculate iqr, upper, lower point for UnitPrice
seventy_fifth = ecommerce_retail['UnitPrice'].quantile(0.75)
twenty_fifth = ecommerce_retail['UnitPrice'].quantile(0.25)
UnitPrice_iqr = seventy_fifth - twenty_fifth
UnitPrice_upper = seventy_fifth + (1.5 * UnitPrice_iqr)
UnitPrice_lower = twenty_fifth - (1.5 * UnitPrice_iqr)

#Calculate iqr, upper, lower point for Quantity
seventy_fifth = ecommerce_retail['Quantity'].quantile(0.75)
twenty_fifth = ecommerce_retail['Quantity'].quantile(0.25)
Quantity_iqr = seventy_fifth - twenty_fifth
Quantity_upper = seventy_fifth + (1.5 * Quantity_iqr)
Quantity_lower = twenty_fifth - (1.5 * Quantity_iqr)

#remove outliers
ecommerce_retail = ecommerce_retail[(ecommerce_retail['UnitPrice'] > UnitPrice_lower) 
                                        & (ecommerce_retail['UnitPrice'] < UnitPrice_upper)
                                    & (ecommerce_retail['Quantity'] > Quantity_lower) 
                                        & (ecommerce_retail['Quantity'] < Quantity_upper)
]

#Recheck outliers in UnitPrice
sns.boxplot(ecommerce_retail, x = 'UnitPrice')  

#Recheck outliers in Quantity
sns.boxplot(ecommerce_retail, x = 'Quantity') 

#Recheck incorrect values, abnormal data, outliers
pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)

print(ecommerce_retail.describe())

#Recheck negative values in UnitPrice and Quantity
print((ecommerce_retail['UnitPrice'] < 0).sum())
print((ecommerce_retail['Quantity'] < 0).sum())

#Count number of duplicates
print("Number of duplicates is: " + str(ecommerce_retail.duplicated().sum()))

#remove duplicates, check duplicates again
ecommerce_retail_removedup = ecommerce_retail.drop_duplicates(
    ["InvoiceNo", "StockCode","InvoiceDate","CustomerID"], keep = 'first'
)

print("Number of duplicates is: " + str(ecommerce_retail_removedup.duplicated().sum()))

pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)

print(ecommerce_retail_removedup)

#Check number of cancelled Invoice
print("Number of cancel Invoice is: " + str((ecommerce_retail_removedup['InvoiceNo'].str.contains("^C")).sum()))

print(ecommerce_retail_removedup.info())

#date for Recency at 31-12-2011
date_R = dt.datetime(2011,12,31)

#Calculate Revenue
ecommerce_retail_removedup['Total_revenue'] = ecommerce_retail_removedup['Quantity'] * ecommerce_retail_removedup['UnitPrice']

#RFM model
RFM_model = ecommerce_retail_removedup.groupby('CustomerID').agg(
    Recency = ('InvoiceDate', lambda x: (date_R - x.max()).days),
    Frequency = ('InvoiceNo', 'count'),
    Monetary = ('Total_revenue', 'sum')
).reset_index()

print(RFM_model)
print(RFM_model.info())

#rank RFM marks
RFM_model= RFM_model.sort_values(
    ['Recency', 'Frequency', 'Monetary'], ascending=(True, False, False)
)

RFM_model['Recency_Score'] = pd.qcut(RFM_model['Recency'], q = 5, labels=[5,4,3,2,1])
RFM_model['Frequency_Score'] = pd.qcut(RFM_model['Frequency'], q = 5, labels=[1,2,3,4,5])
RFM_model['Monetary_Score'] = pd.qcut(RFM_model['Monetary'], q = 5, labels=[1,2,3,4,5])

#Calculate RFM Score
RFM_model['RFM_Score'] = (
    RFM_model['Recency_Score'].astype(str) + RFM_model['Frequency_Score'].astype(str) + RFM_model['Monetary_Score'].astype(str)
)
RFM_model['RFM_Score'] = RFM_model['RFM_Score'].astype(int)

pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)
print(RFM_model)
print(RFM_model.info())

#Check outliers in Recency
sns.boxplot(RFM_model, x = 'Recency')

#Check outliers in Frequency
sns.boxplot(RFM_model, x = 'Frequency')

#Check outliers in Monetary
sns.boxplot(RFM_model, x = 'Monetary')

#Setup IQR, upper, lower point for Recency
Seventy_fifth = RFM_model['Recency'].quantile(0.75)
Twenty_fifth = RFM_model['Recency'].quantile(0.25)
Recency_iqr = Seventy_fifth - Twenty_fifth
Recency_upper = Seventy_fifth + (1.5 * Recency_iqr)
Recency_lower = Twenty_fifth - (1.5 * Recency_iqr)

#Setup IQR, upper, lower point for Frequency
Seventy_fifth = RFM_model['Frequency'].quantile(0.75)
Twenty_fifth = RFM_model['Frequency'].quantile(0.25)
Frequency_iqr = Seventy_fifth - Twenty_fifth
Frequency_upper = Seventy_fifth + (1.5 * Frequency_iqr)
Frequency_lower = Twenty_fifth - (1.5 * Frequency_iqr)

#Setup IQR, upper, lower point for Monetary
Seventy_fifth = RFM_model['Monetary'].quantile(0.75)
Twenty_fifth = RFM_model['Monetary'].quantile(0.25)
Monetary_iqr = Seventy_fifth - Twenty_fifth
Monetary_upper = Seventy_fifth + (1.5 * Monetary_iqr)
Monetary_lower = Twenty_fifth - (1.5 * Monetary_iqr)

#Remove Outliers
RFM_model_remove_outliners = RFM_model[(RFM_model['Recency'] > Recency_lower) 
                                            & (RFM_model['Recency'] < Recency_upper)
                                       & (RFM_model['Frequency'] > Frequency_lower) 
                                            & (RFM_model['Frequency'] < Frequency_upper)
                                       & (RFM_model['Monetary'] > Monetary_lower) 
                                             & (RFM_model['Monetary'] < Monetary_upper)
]

#Recheck outliers in Recency
sns.boxplot(RFM_model_remove_outliners, x = 'Recency')

#Recheck outliers in Frequency
sns.boxplot(RFM_model_remove_outliners, x = 'Frequency')

#Recheck outliers in Monetary
sns.boxplot(RFM_model_remove_outliners, x = 'Monetary')

#import excel file, sheet 2
# Segmentation = pd.read_excel(
#     r'C:\\Users\\Admin\\Desktop\\DA books\\DA lesson\\Python\\Final_project_RFM\\ecommerce retail.xlsx', sheet_name = 'Segmentation'
# )
# print(Segmentation)

# Split values
Segmentation['RFM Score'] = Segmentation['RFM Score'].str.split(',')
Segmentation_sepe = Segmentation.explode('RFM Score').reset_index(drop = True)
Segmentation_sepe.rename(columns= {'RFM Score': 'RFM_Score'}, inplace= True)

Segmentation_sepe['RFM_Score'] = Segmentation_sepe['RFM_Score'].astype(int)

print(Segmentation_sepe)

RFM_Seg = pd.merge(
    RFM_model_remove_outliners, Segmentation_sepe, on= 'RFM_Score', how = 'left'
)

pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)

print(RFM_Seg)

Cus_Seg = RFM_Seg.merge(
    ecommerce_retail_removedup[['CustomerID','Country']], on ='CustomerID', how = 'inner'
)
Cus_Seg = Cus_Seg.drop_duplicates()
pd.set_option('display.max_columns', None)  
pd.set_option('display.width', 1000)
print(Cus_Seg)

plt.figure(figsize=(10, 6))
sns.set_palette("coolwarm")
sns.set_style("dark")
g= sns.catplot(x = "Segment", y ="Recency", data = RFM_Seg, kind = "bar", palette= "coolwarm", ci= None)
g.fig.suptitle("Recency Value by Segmentation", y = 1.1)
g.set(xlabel= "Segmentation",
      ylabel= "Recency")
plt.xticks(rotation = 90)
plt.show()

plt.figure(figsize=(10, 6))
sns.set_palette("coolwarm")
sns.set_style("dark")
g= sns.catplot(x = "Segment", y ="Frequency", data = RFM_Seg, kind = "bar", palette= "coolwarm", ci= None)
g.fig.suptitle("Frequency Value by Segmentation", y = 1.1)
g.set(xlabel= "Segmentation",
      ylabel= "Frequency")
plt.xticks(rotation = 90)
plt.show()

plt.figure(figsize=(10, 6))
sns.set_palette("coolwarm")
sns.set_style("dark")
g= sns.catplot(x = "Segment", y ="Monetary", data = RFM_Seg, kind = "bar", palette= "coolwarm", ci= None)
g.fig.suptitle("Monetary Value by Segmentation", y = 1.1)
g.set(xlabel= "Segmentation",
      ylabel= "Monetary")
plt.xticks(rotation = 90)
plt.show()

colnames = ['Recency', 'Frequency', 'Monetary']

for col in colnames:
    fig, ax = plt.subplots(figsize=(12,3))
    sns.distplot(RFM_Seg[col])
    ax.set_title('Distribution of %s' % col)
    plt.show()

# As can be seen from the chart that :
# -Hibernating Customers occupy the highest density in this plot (22.09%), followed by the number of customer who are Potential Loyalist (13.7%), showing a potential chance for growth of company in loyal segment. While Loyay Customers, New Customers and Champions share nearly similar proportion (about 9.3%, 8.8% and 10.53% respectively), this group is important for enhancing the company'sustainability and growth. In contrast the number of At Risk customers have a quite high density (about 11.2%).

# -Although the number of valuable customers, including potential, loyal and new customers help for steady growth , we need to pay attention on At Risk and Need Attention Customer for posing some approriate stategies to prevent increasing churn rate.

# -Potential Loyalist have high frequency (mean at 46) while have the low recency(mean at 50) =>  paying attention for converting them a loyal customer could be a priority . Besides, Hibernating still be at high rate, frequency (mean at 18) , Recency (nearly 164) indicate that they used to be a loyal or valuable customers but recently they have no any transaction => be focus on.

# => Current state of company still be sustainable, the rate of valuable customers occupy at high density to ensure the growth of company, while the number of At risk customers show a high proportion that could lead to increasing churn rate, followed by other group such as need attention or hibernating customer , having negative effect for trustworthy of company 


# #Suggest for Marketing and Sales team retail model, should we focus on R, F or M metric in RFM model:
# For retail company, R and F should be prioritized as key metrics for them. These metris provide insights how recently customers have returned and  how often they buy in order to have appropriate campaigns for company.

# Firstly, for Potential Loyalist, we can see that they have high frequency over long-term period with low recency , we should make a survey for evaluating their satisfaction, give more promotion code, discount voucher with free shipping fee to increase their loyalness, encourage them to buy more products for converting them loyal customers.

# Secondly, for At Risk Customers, we need to analyse several reasons that lead to churness or being less active. Then,  proposing promotion program or customer appprciation with buy one get one event in order to engage them.

# Thirdly, Hibernating customers had high frequency  long-term period in the past, but recently they have no activity. We should have customer apppreciation program for welcoming back, promotion code ,reactivation discount and reviewing customer care , special gifts in special days like New Year or Christmas day for getting customers back.