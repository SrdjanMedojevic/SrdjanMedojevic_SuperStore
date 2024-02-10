#!/usr/bin/env python
# coding: utf-8

# # SuperStoreUS
# ## Data analysis

# ## Preparing and Analysing Data

# In[29]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
get_ipython().run_line_magic('matplotlib', 'inline')
import seaborn as sns


# In[2]:


#sheet 1
df1 = pd.read_excel(r"C:\Users\Korisnik\Desktop\vega it\SuperStoreUS.xlsx", sheet_name="Orders")
print(df1)


# In[3]:


#sheet 2
df2 = pd.read_excel(r"C:\Users\Korisnik\Desktop\vega it\SuperStoreUS.xlsx", sheet_name="Returns")
print(df2)


# In[4]:


#sheet 3
df3 = pd.read_excel(r"C:\Users\Korisnik\Desktop\vega it\SuperStoreUS.xlsx", sheet_name="Users")
print(df3)


# In[5]:


#checking for type
df1.dtypes


# In[6]:


#checking for nulls
df1.isnull().any()


# In[7]:


#fixing NaN
df1 [df1['Product Base Margin'].isna()==1]


# In[8]:


#fixing NaN
df1['Product Base Margin'] = df1['Product Base Margin'].fillna(0)
df1.isnull().any()


# In[9]:


#seraching for duplicates
df1.duplicated()


# In[10]:


#info about df
df1.info()


# In[11]:


#merging df1 and df2 (returned)
df4=df1.merge(df2, on = 'Order ID', how='inner')
df4


# In[12]:


#merging df1 and df3 (matching manager with regions form df1)
df5=df1.merge(df3, on = 'Region', how='left')
df5


# In[13]:


#description of df1
df1.describe()


# In[14]:


df5.nunique()


# ## Data analysis and visualization

# In[53]:


#Count of Shipments by Ship mode
pivot_df1 = df5.pivot_table(index='Ship Mode', columns='Region', values='Shipping Cost', aggfunc='count')
pivot_df1
pivot_df1.T.plot(marker='o', linestyle='-')
plt.title('Count of Shipments by Ship Mode and Region')
plt.xlabel('Region')
plt.ylabel('Count of Shipments')
plt.legend(title='Ship Mode', bbox_to_anchor=(0.25, 0.25))
plt.show()


# In[54]:


#The pivot code in combination with the heatmap clearly shows us how much more expensive it actually is to transport goods by truck than by plane. 
pivot_df = df5.pivot_table(index='Ship Mode', columns='Region', values='Shipping Cost', aggfunc='mean')
pivot_df
plt.figure(figsize=(8, 4))
sns.heatmap(pivot_df, cmap='viridis', annot=True, fmt='.2f', cbar_kws={'label': 'Average Shipping Cost'})
plt.title('Average Shipping Cost by Ship Mode and Region')
plt.show()


# In[77]:


#top 10 largest profits by state
top_10_largest_profit = df5.nlargest(10, 'Profit')
print("Top 10 largest Profits by Order ID:")
print(top_10_largest_profit[['Order ID', 'Profit', 'State or Province']])


# In[73]:


#top 10 lowest profits by state
top_10_lowest_profits = df5.nsmallest(10, 'Profit')
print("Top 10 Lowest Profits by Order ID:")
print(top_10_lowest_profits[['Order ID', 'Profit', 'State or Province']])


# In[62]:


#On this graph, we can see that the countries with the highest incomes are also the countries with the largest number of people.
#North Carolina and Montana have lowest profits in top 10 which puts them at the bottom. 
pivot_df = df5.pivot_table(index='State or Province', columns='Country', values='Profit', aggfunc='sum')
pivot_df
import plotly.express as px
pivot_df_long = pivot_df.reset_index().melt(id_vars='State or Province', var_name='Country', value_name='Profit')
fig = px.bar(pivot_df_long,
             x='State or Province',
             y='Profit',
             color='Country',
             title='Aggregated Profit by State and Country',
             labels={'Profit': 'Total Profit'},
             height=600)
fig.show()


# ### What sells the most

# In[178]:


#Product Sub-Category
##The most expensive was furniture (Chairs & Chairmats), but office suppliese (Binders and Binder Accessories) was bought the most.
product_category_stats = df5.groupby('Product Category')['Sales'].agg(['max', 'min', 'mean', 'count']).reset_index()
product_category_stats


# In[177]:


#Product Sub-Category
product_category_stats = df5.groupby(['Product Category', 'Product Sub-Category'])['Sales'].agg(['max', 'min', 'mean', 'count']).reset_index()
product_category_stats1


# In[109]:


pivot_df5 = df5.pivot_table(index='Product Category', columns='Product Sub-Category', values='Profit', aggfunc='sum')
pivot_df5
c_pivot_df5 = pivot_df.style.highlight_max(axis=0, color='brown').highlight_min(axis=0, color='green')
c_pivot_df5 


# In[123]:


profit_per_category = df5.groupby('Product Category')['Profit'].sum()
fig, axs = plt.subplots(1, 1, figsize=(6, 3))
axs.pie(profit_per_category, labels=profit_per_category.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
axs.set_title('Total Profit per Product Category')
plt.tight_layout()
plt.show()


# In[128]:


profit_per_subcategory = df5.groupby('Product Sub-Category')['Profit'].sum()
plt.figure(figsize=(6, 3))
profit_per_category.plot(kind='bar', color='skyblue')
plt.title('Total Profit per Product Sub-Category')
plt.xlabel('Product Sub-Category')
plt.ylabel('Total Profit')
plt.xticks(rotation=45, ha='right')  # Rotate x-axis labels for better readability
plt.show()


# In[168]:


#Manager_Profit
Manager_Profit = df5.groupby('Manager')['Profit'].agg(['max', 'min', 'mean','count']).reset_index()
Manager_Profit


# In[169]:


#Manager Erin is the most successful
fig, axs = plt.subplots(2, 1, figsize=(7, 5.5))
axs[0].bar(Manager_Profit['Manager'], Manager_Profit['max'], label='Max Profit', alpha=0.7)
axs[0].bar(Manager_Profit['Manager'], Manager_Profit['min'], label='Min Profit', alpha=0.7)
axs[0].set_ylabel('Profit')
axs[0].set_title('Max, Min for Each Manager')
axs[0].legend()
axs[1].bar(Manager_Profit['Manager'], Manager_Profit['count'], color='pink', label='Record Count')
axs[1].set_ylabel('Record Count')
axs[1].set_title('Number of Records for Each Manager')
axs[1].legend()
plt.tight_layout()
plt.show()


# In[140]:


#From the table that describes the correlation, we can see that the discount and shipping cost affect the profit.
columns_to_correlate = ['Profit', 'Shipping Cost', 'Discount']
correlation = df5[columns_to_correlate].corr()
print("Correlation:")
print(correlation)


# In[155]:


#We can see that the price of shipping is most affected by jambo packaging.
plt.scatter(df5['Product Container'], df5['Shipping Cost'], alpha=0.5, color='blue')
plt.title('Scatter Plot: Product Container vs Shipping Cost')
plt.xlabel('Product Container')
plt.ylabel('Shipping Cost')
plt.xticks(rotation=45)
plt.show()


# In[160]:


#top_15_largest_shipping_cost
top_15_largest_shipping_cost = df5.nlargest(15, 'Shipping Cost')
print("Top_15_largest_shipping_cost by Order ID:")
print(top_15_largest_profit[['Product Container', 'Shipping Cost']])

