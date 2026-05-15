import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

# ==========================================================
# SALES INSIGHTS PIPELINE
# Answers the following questions:
# 1. What is total revenue?
# 2. What is month-over-month growth?
# 3. What is year-over-year growth?
# 4. What is the most ordered time-window?
# 5. What is customer lifetime value?
# 6. What is customer churn rate?
# 7. What is average order value?
# 8. What is sales per region/category/product?
# 9. Which products have declining sales?
# 10. Product vs month evaluation: which product is ordered most?
# ==========================================================

# Create output folders
Path("output/charts").mkdir(parents=True, exist_ok=True)
Path("output/reports").mkdir(parents=True, exist_ok=True)

# ==========================================================
# LOAD DATASET
# Excel file should contain columns:
# Date, Order ID, Customer ID, Product, Category,
# Region, Quantity, Price
# ==========================================================
print("Loading dataset...")
df = pd.read_excel("S3-4_Capstone_Task_Pick_dataset_define_3_questions_set_up_skeleton_notebook (1).xlsx")

# Clean column names
df.columns = df.columns.str.strip()

# Convert Date column to datetime
df["Date"] = pd.to_datetime(df["Date"])

# Create calculated columns
df["Revenue"] = df["Quantity"] * df["Price"]
df["Month"] = df["Date"].dt.to_period("M").astype(str)
df["Year"] = df["Date"].dt.year
def format_currency(series):
    return series.apply(lambda x: f"Rs {x:,.2f}")

# Sort by date for proper time-series analysis
df = df.sort_values("Date")

print("\n========== SALES INSIGHTS ==========")

# ==========================================================
# 1. TOTAL REVENUE
# ==========================================================
total_revenue = df["Revenue"].sum()
print(f"\n1. Total Revenue: Rs {total_revenue:,.2f}")

# ==========================================================
# 2. AVERAGE ORDER VALUE
# ==========================================================
# Average revenue per unique order
order_revenue = df.groupby("Order ID")["Revenue"].sum()
average_order_value = order_revenue.mean()
print(f"\n2. Average Order Value: Rs {average_order_value:,.2f}")

# ==========================================================
# 3. MONTH-OVER-MONTH GROWTH
# ==========================================================
monthly_sales = df.groupby("Month")["Revenue"].sum().sort_index()
mom_growth = (monthly_sales.pct_change() * 100).fillna(0)

print("\n3. Month-over-Month Growth (%):")
print(mom_growth.round(2))

# ==========================================================
# 4. YEAR-OVER-YEAR GROWTH
# ==========================================================
yearly_sales = df.groupby("Year")["Revenue"].sum().sort_index()
yoy_growth = (yearly_sales.pct_change() * 100).fillna(0)

print("\n4. Year-over-Year Growth (%):")
print(yoy_growth.round(2))

# ==========================================================
# 5. MOST ORDERED TIME-WINDOW
# ==========================================================
# Month with the highest number of orders
orders_per_month = df.groupby("Month")["Order ID"].nunique()
most_ordered_time_window = orders_per_month.idxmax()

print(f"\n5. Most Ordered Time Window: {most_ordered_time_window}")

# ==========================================================
# 6. CUSTOMER LIFETIME VALUE (CLV)
# ==========================================================
customer_lifetime_value = (
    df.groupby("Customer ID")["Revenue"]
    .sum()
    .sort_values(ascending=False)
)

print("\n6. Customer Lifetime Value (Top 10 Customers):")
print(format_currency(customer_lifetime_value.head(10)))

# ==========================================================
# 7. CUSTOMER CHURN RATE
# ==========================================================
# Customers who purchased in the last 30 days are considered active
latest_date = df["Date"].max()
cutoff_date = latest_date - pd.Timedelta(days=30)

total_customers = df["Customer ID"].nunique()
active_customers = df.loc[
    df["Date"] >= cutoff_date, "Customer ID"
].nunique()

churn_rate = ((total_customers - active_customers) / total_customers) * 100

print(f"\n7. Customer Churn Rate: {churn_rate:.2f}%")

# ==========================================================
# 8. SALES PER REGION
# ==========================================================
region_sales = (
    df.groupby("Region")["Revenue"]
    .sum()
    .sort_values(ascending=False)
)
print("\n8. Sales per Region:")
print(format_currency(region_sales))

# ==========================================================
# 9. SALES PER CATEGORY
# ==========================================================
category_sales = (
    df.groupby("Category")["Revenue"]
    .sum()
    .sort_values(ascending=False)
)

print("\n9. Sales per Category:")
print(format_currency(category_sales))
# ==========================================================
# 10. SALES PER PRODUCT
# ==========================================================
product_sales = (
    df.groupby("Product")["Revenue"]
    .sum()
    .sort_values(ascending=False)
)

print("\n10. Sales per Product:")
print(format_currency(product_sales))

# ==========================================================
# 11. PRODUCTS WITH DECLINING SALES
# ==========================================================
# Compare first month sales vs last month sales
product_month_sales = (
    df.groupby(["Month", "Product"])["Revenue"]
    .sum()
    .unstack(fill_value=0)
    .sort_index()
)

declining_products = (
    product_month_sales.iloc[-1] - product_month_sales.iloc[0]
)

declining_products = declining_products[declining_products < 0].sort_values()

print("\n11. Products with Declining Sales (Last Month Revenue - First Month Revenue):")
if declining_products.empty:
    print("No products with declining sales.")
else:
    print(format_currency(declining_products).to_string())
# ==========================================================
# 12. PRODUCT VS MONTH EVALUATION
# Which product is ordered the most each month?
# ==========================================================
product_vs_month = (
    df.groupby(["Month", "Product"])["Quantity"]
    .sum()
    .reset_index()
)

most_ordered_product_by_month = (
    product_vs_month.loc[
        product_vs_month.groupby("Month")["Quantity"].idxmax()
    ]
    .sort_values("Month")
)

print("\n12. Most Ordered Product by Month:")
print(most_ordered_product_by_month.to_string(index=False))
# Overall most ordered product
# Overall most ordered product
overall_most_ordered_product = (
    df.groupby("Product")["Quantity"]
    .sum()
    .sort_values(ascending=False)
)

max_qty = overall_most_ordered_product.max()
top_products = overall_most_ordered_product[
    overall_most_ordered_product == max_qty
]

print("\nOverall Most Ordered Product(s):")
for product, qty in top_products.items():
    print(f"{product} ({qty} units)")
# ==========================================================
# SAVE REPORTS
# ==========================================================
print("\nSaving reports...")

monthly_sales.to_csv("output/reports/monthly_sales.csv")
mom_growth.to_csv("output/reports/mom_growth.csv")
yearly_sales.to_csv("output/reports/yearly_sales.csv")
yoy_growth.to_csv("output/reports/yoy_growth.csv")
customer_lifetime_value.to_csv("output/reports/customer_lifetime_value.csv")
region_sales.to_csv("output/reports/region_sales.csv")
category_sales.to_csv("output/reports/category_sales.csv")
product_sales.to_csv("output/reports/product_sales.csv")
declining_products.to_csv("output/reports/declining_products.csv")
most_ordered_product_by_month.to_csv(
    "output/reports/most_ordered_product_by_month.csv",
    index=False
)

# Save all summaries into one Excel workbook
with pd.ExcelWriter("output/reports/sales_insights_summary.xlsx") as writer:
    monthly_sales.to_excel(writer, sheet_name="Monthly Sales")
    mom_growth.to_excel(writer, sheet_name="MoM Growth")
    yearly_sales.to_excel(writer, sheet_name="Yearly Sales")
    yoy_growth.to_excel(writer, sheet_name="YoY Growth")
    customer_lifetime_value.to_excel(writer, sheet_name="Customer CLV")
    region_sales.to_excel(writer, sheet_name="Region Sales")
    category_sales.to_excel(writer, sheet_name="Category Sales")
    product_sales.to_excel(writer, sheet_name="Product Sales")
    declining_products.to_excel(writer, sheet_name="Declining Products")
    most_ordered_product_by_month.to_excel(
        writer,
        sheet_name="Most Ordered Product",
        index=False
    )

# ==========================================================
# VISUALIZATIONS
# ==========================================================
print("Generating charts...")
sns.set_style("whitegrid")

# 1. Product Sales Bar Chart
plt.figure(figsize=(12, 6))
sns.barplot(x=product_sales.index, y=product_sales.values)
plt.title("Total Revenue by Product")
plt.xlabel("Product")
plt.ylabel("Revenue")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/charts/product_sales.png")
plt.close()

# 2. Monthly Sales Line Chart
plt.figure(figsize=(12, 6))
plt.plot(monthly_sales.index, monthly_sales.values, marker="o")
plt.title("Monthly Total Revenue")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/charts/monthly_sales.png")
plt.close()

# 3. Revenue Distribution Histogram
plt.figure(figsize=(12, 6))
plt.hist(df["Revenue"], bins=20, edgecolor="black")
plt.title("Revenue Distribution")
plt.xlabel("Revenue")
plt.ylabel("Frequency")
plt.tight_layout()
plt.savefig("output/charts/revenue_distribution.png")
plt.close()

# 4. Region Sales Chart
plt.figure(figsize=(12, 6))
sns.barplot(x=region_sales.index, y=region_sales.values)
plt.title("Revenue by Region")
plt.xlabel("Region")
plt.ylabel("Revenue")
plt.tight_layout()
plt.savefig("output/charts/region_sales.png")
plt.close()

# 5. Category Sales Chart
plt.figure(figsize=(12, 6))
sns.barplot(x=category_sales.index, y=category_sales.values)
plt.title("Revenue by Category")
plt.xlabel("Category")
plt.ylabel("Revenue")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/charts/category_sales.png")
plt.close()

# 6. Month-over-Month Growth Chart
plt.figure(figsize=(12, 6))
plt.plot(mom_growth.index, mom_growth.values, marker="o")
plt.title("Month-over-Month Growth (%)")
plt.xlabel("Month")
plt.ylabel("Growth %")
plt.axhline(0, linestyle="--")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("output/charts/mom_growth.png")
plt.close()

print("\nAll reports and charts saved successfully!")
print("Reports saved in: output/reports/")
print("Charts saved in: output/charts/")