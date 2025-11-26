📊 Regional Sales Data Analysis (Python EDA)

This project performs a comprehensive Exploratory Data Analysis (EDA) on Acme Co.’s multi-sheet Regional Sales Dataset (2014–2018).
The goal is to uncover revenue drivers, profitability insights, seasonal trends, channel performance, and regional strengths, while aligning actual performance against 2017 budget targets.

🚀 Project Overview

Using Python (Pandas, Matplotlib, Seaborn), this analysis:

Consolidates data from multiple Excel sheets (Sales Orders, Products, Customers, Regions, Budgets, etc.)

Cleans, transforms, and merges datasets into a single master dataset

Calculates key business metrics like total cost, profit, profit margin, AOV, product profitability, and more

Identifies top-performing products, high-margin items, winning regions, and dominant sales channels

Visualizes trends such as:

Monthly sales (2014–2018)

Product revenue leaders

Channel-level contribution

Profit margin distribution

Unit price variations

Regional performance

The insights contribute to pricing, promotion, and market expansion decisions, and are designed to support a future Power BI dashboard.

📁 Dataset Structure

The Excel file contains these sheets:

Sheet Name	Description
Sales Orders	Main order-level data
Customers	Customer demographic & channel data
Products	Product metadata
Regions	US regions mapping
State Regions	State → Region mapping
2017 Budgets	Budgeted sales by product
🧹 Data Cleaning & Preparation

The notebook performs:

Header correction (State Regions sheet)

Joining all sheets into a unified dataset

Type conversions (datetime, numeric conversions)

Missing value handling

Feature engineering:

total_cost = order_quantity * cost

profit = revenue - total_cost

profit_margin_pct

order_month

Month names and numbers for seasonal analysis

Filtering out incomplete 2018 data when needed

Ensuring budgets apply only to 2017 actuals

📈 Key Analyses & Visualizations
✔ Monthly Sales Trend

Trendline from 2014–2017

Seasonal patterns (peaks around Q4)

✔ Top 10 Products by Revenue

Identifies highest-earning products

✔ Top 10 Products by Profit Margin

Highlights most profitable items

✔ Channel Performance

Pie chart of total revenue contribution (Retail, Distributor, Online, etc.)

✔ AOV (Average Order Value) Distribution

Histogram showing spending patterns per order

✔ Unit Price vs Profit Margin

Scatterplot showing correlation between pricing and profitability

✔ Product-wise Unit Price Distribution

Boxplot of price variations across products

✔ Sales by Region

Bar chart of US regions ranked by revenue

🛠 Tech Stack

Python

Pandas

NumPy

Matplotlib

Seaborn

Jupyter Notebook / Google Colab

How to Run
git clone https://github.com/miteshh16/Sales-Forecasting-Performance-Insights.git
cd code.py

Insights Generated

Identified revenue concentration among a few products → risk for business

Regions with highest growth potential highlighted

High-margin vs low-margin product categories

Seasonal peaks helpful for demand planning and targeted promotions

Channel-wise strengths to inform marketing budget allocation

Pricing–profitability relationship visualized to avoid underpricing
