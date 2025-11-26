# ============================
#  SALES ANALYSIS PROJECT
#  Dataset: Regional Sales Dataset.xlsx
#  Author: Your Name
#  ============================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import FuncFormatter
import warnings
warnings.filterwarnings('ignore')

# ======================================
# 1. LOAD DATASET (ALL SHEETS)
# ======================================
sheets = pd.read_excel("Regional Sales Dataset.xlsx", sheet_name=None)

df_sales = sheets["Sales Orders"]
df_customers = sheets["Customers"]
df_products = sheets["Products"]
df_regions = sheets["Regions"]
df_state_reg = sheets["State Regions"]
df_budgets = sheets["2017 Budgets"]

print(f"df_sales shape: {df_sales.shape}")
print(f"df_customers shape: {df_customers.shape}")
print(f"df_products shape: {df_products.shape}")
print(f"df_regions shape: {df_regions.shape}")
print(f"df_state_reg shape: {df_state_reg.shape}")
print(f"df_budgets shape: {df_budgets.shape}")

# ======================================
# 2. CLEAN STATE REGION SHEET
# ======================================
new_header = df_state_reg.iloc[0]
df_state_reg.columns = new_header
df_state_reg = df_state_reg[1:].reset_index(drop=True)

# ======================================
# 3. MERGE ALL DATASETS
# ======================================
df = df_sales.merge(df_customers, how="left",
                    left_on="Customer Name Index",
                    right_on="Customer Index")

df = df.merge(df_products, how="left",
              left_on="Product Description Index",
              right_on="Index")

df = df.merge(df_regions, how="left",
              left_on="Delivery Region Index",
              right_on="id")

df = df.merge(df_state_reg[["State Code", "Region"]], how="left",
              left_on="state_code",
              right_on="State Code")

df = df.merge(df_budgets, how="left", on="Product Name")

cols_to_drop = ["Customer Index", "Index", "id", "State Code"]
df.drop(columns=cols_to_drop, errors="ignore", inplace=True)

# ======================================
# 4. CLEAN COLUMN NAMES
# ======================================
df.columns = df.columns.str.lower()

# Keep necessary columns
cols_to_keep = [
    'ordernumber', 'orderdate', 'customer names', 'channel',
    'product name', 'order quantity', 'unit price', 'line total',
    'total unit cost', 'county', 'state', 'region',
    'latitude', 'longitude', '2017 budgets'
]

df = df[cols_to_keep]

df = df.rename(columns={
    'ordernumber': 'order_number',
    'orderdate': 'order_date',
    'customer names': 'customer_name',
    'product name': 'product_name',
    'order quantity': 'order_quantity',
    'unit price': 'unit_price',
    'line total': 'revenue',
    'total unit cost': 'cost',
    '2017 budgets': 'budget',
    'latitude': 'lat',
    'longitude': 'lon'
})

# ======================================
# 5. BUDGET CLEANUP (ONLY FOR 2017)
# ======================================
df.loc[df['order_date'].dt.year != 2017, 'budget'] = pd.NA

# ======================================
# 6. FEATURE ENGINEERING
# ======================================
df['total_cost'] = df['order_quantity'] * df['cost']
df['profit'] = df['revenue'] - df['total_cost']
df['profit_margin_pct'] = (df['profit'] / df['revenue']) * 100
df['order_month'] = df['order_date'].dt.to_period('M')

# ======================================
# 7. MONTHLY SALES TREND (ALL YEARS)
# ======================================
monthly_sales = df.groupby("order_month")["revenue"].sum()

plt.figure(figsize=(15, 4))
monthly_sales.plot(marker="o", color="navy")
plt.gca().yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x/1e6:.1f}M"))
plt.title("Monthly Sales Analysis (All Years)")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# ======================================
# 8. MONTHLY SALES EXCLUDING 2018
# ======================================
df_no2018 = df[df['order_date'].dt.year != 2018]
df_no2018['month_num'] = df_no2018['order_date'].dt.month
df_no2018['month_name'] = df_no2018['order_date'].dt.strftime('%B')

monthly_sales = (
    df_no2018.groupby(['month_num', 'month_name'])['revenue']
    .sum()
    .sort_index()
)

plt.figure(figsize=(13, 4))
plt.plot(monthly_sales.index.get_level_values(1), monthly_sales.values,
         marker="o", color="navy")

plt.gca().yaxis.set_major_formatter(FuncFormatter(lambda x, pos: f"{x/1e6:.1f}M"))
plt.title("Monthly Sales Trend (Excluding 2018)")
plt.xlabel("Month")
plt.ylabel("Revenue (Millions)")
plt.xticks(rotation=35)
plt.tight_layout()
plt.show()

# ======================================
# 9. TOP 10 PRODUCTS BY REVENUE
# ======================================
top_prod = df.groupby("product_name")["revenue"].sum().nlargest(10) / 1_000_000

plt.figure(figsize=(9, 4))
sns.barplot(x=top_prod.values, y=top_prod.index, palette="viridis")
plt.title("Top 10 Products by Revenue (in Millions)")
plt.xlabel("Revenue (Millions)")
plt.ylabel("Product Name")
plt.tight_layout()
plt.show()

# ======================================
# 10. TOP 10 PRODUCTS BY AVG PROFIT
# ======================================
top_margin = df.groupby("product_name")["profit"].mean().nlargest(10)

plt.figure(figsize=(9, 4))
sns.barplot(x=top_margin.values, y=top_margin.index, palette="viridis")
plt.title("Top 10 Products by Avg Profit")
plt.xlabel("Average Profit (USD)")
plt.ylabel("Product Name")
plt.tight_layout()
plt.show()

# ======================================
# 11. SALES BY CHANNEL
# ======================================
channel_sales = df.groupby("channel")["revenue"].sum()

plt.figure(figsize=(5, 5))
plt.pie(channel_sales.values, labels=channel_sales.index,
        autopct='%1.1f%%', startangle=140,
        colors=sns.color_palette("coolwarm"))
plt.title("Sales by Channel")
plt.tight_layout()
plt.show()

# ======================================
# 12. AVERAGE ORDER VALUE (AOV)
# ======================================
aov = df.groupby("order_number")["revenue"].sum()

plt.figure(figsize=(12, 4))
plt.hist(aov, bins=50, color="navy", edgecolor="black")
plt.title("Average Order Value Distribution")
plt.xlabel("Order Value (USD)")
plt.ylabel("Frequency")
plt.tight_layout()
plt.show()

# ======================================
# 13. PROFIT MARGIN VS UNIT PRICE
# ======================================
plt.figure(figsize=(6, 4))
plt.scatter(df["unit_price"], df["profit_margin_pct"], alpha=0.6, color="green")
plt.title("Profit Margin % vs Unit Price")
plt.xlabel("Unit Price (USD)")
plt.ylabel("Profit Margin (%)")
plt.tight_layout()
plt.show()

# ======================================
# 14. UNIT PRICE DISTRIBUTION PER PRODUCT
# ======================================
plt.figure(figsize=(12, 4))
sns.boxplot(data=df, x="product_name", y="unit_price", color="g")
plt.title("Unit Price by Product")
plt.xlabel("Product")
plt.ylabel("Unit Price")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.show()

# ======================================
# 15. SALES BY REGION
# ======================================
region_sales = df.groupby("region")["revenue"].sum().sort_values(ascending=False)

plt.figure(figsize=(10, 4))
sns.barplot(x=region_sales.values, y=region_sales.index, palette="Greens_r")
plt.title("Total Sales by US Region")
plt.xlabel("Revenue (Millions USD)")
plt.ylabel("Region")
plt.tight_layout()
plt.show()
