from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "Sales_Data.xlsx"
OUT = HERE / "output"
OUT.mkdir(exist_ok=True)

sales = pd.read_excel(SOURCE, sheet_name="Sales Data")
sales.columns = sales.columns.str.strip()
sales = sales.drop_duplicates().copy()
sales["Purchase Date"] = pd.to_datetime(sales["Purchase Date"], errors="coerce")
for col in ["Quantity", "Unit Price", "Total Price", "Discount Applied (%)", "Discount Amount", "Final Price"]:
    sales[col] = pd.to_numeric(sales[col], errors="coerce")
sales = sales.dropna(subset=["Order ID", "Purchase Date", "Product ID", "Final Price"])
sales.to_csv(OUT / "clean_sales_data.csv", index=False)

products = pd.read_excel(SOURCE, sheet_name="Products")
products.columns = products.columns.str.strip()

product_sales = sales.groupby("Product ID", as_index=False).agg(
    units_sold=("Quantity", "sum"),
    revenue=("Final Price", "sum"),
    gross_sales=("Total Price", "sum"),
    discount_amount=("Discount Amount", "sum"),
)
product_sales = product_sales.merge(products[["Product ID", "Product Name", "Category"]], on="Product ID", how="left")
product_sales.sort_values("revenue", ascending=False).to_csv(OUT / "product_summary.csv", index=False)

category_sales = product_sales.groupby("Category", as_index=False).agg(
    units_sold=("units_sold", "sum"),
    revenue=("revenue", "sum"),
    gross_sales=("gross_sales", "sum"),
).sort_values("revenue", ascending=False)
category_sales.to_csv(OUT / "category_summary.csv", index=False)

location_sales = sales.groupby("Store Location", as_index=False).agg(
    orders=("Order ID", "nunique"),
    units_sold=("Quantity", "sum"),
    revenue=("Final Price", "sum"),
).sort_values("revenue", ascending=False)
location_sales.to_csv(OUT / "location_summary.csv", index=False)

monthly = sales.assign(Month=sales["Purchase Date"].dt.to_period("M").astype(str)).groupby("Month", as_index=False).agg(
    orders=("Order ID", "nunique"),
    units_sold=("Quantity", "sum"),
    revenue=("Final Price", "sum"),
    gross_sales=("Total Price", "sum"),
    discount_amount=("Discount Amount", "sum"),
)
monthly.to_csv(OUT / "monthly_sales.csv", index=False)

top_unit_product = product_sales.sort_values("units_sold", ascending=False).iloc[0]["Product Name"]
top_revenue_product = product_sales.sort_values("revenue", ascending=False).iloc[0]["Product Name"]
top_location = location_sales.iloc[0]["Store Location"]

kpis = {
    "rows_cleaned": len(sales),
    "unique_orders": sales["Order ID"].nunique(),
    "units_sold": sales["Quantity"].sum(),
    "gross_sales_before_discount": sales["Total Price"].sum(),
    "discount_amount": sales["Discount Amount"].sum(),
    "net_revenue_final_price": sales["Final Price"].sum(),
    "top_product_by_units": top_unit_product,
    "top_product_by_revenue": top_revenue_product,
    "top_location_by_revenue": top_location,
    "date_start": sales["Purchase Date"].min().date().isoformat(),
    "date_end": sales["Purchase Date"].max().date().isoformat(),
}
pd.Series(kpis, name="value").to_csv(OUT / "kpis.csv", header=True)

plt.style.use("seaborn-v0_8-whitegrid")
fig, axes = plt.subplots(2, 2, figsize=(14, 10))

axes[0, 0].bar(monthly["Month"], monthly["revenue"], color="#176B87")
axes[0, 0].set_title("Monthly Net Revenue", fontsize=12, fontweight="bold", pad=10)
axes[0, 0].set_xlabel("Month")
axes[0, 0].set_ylabel("Net Revenue")
axes[0, 0].tick_params(axis="x", rotation=60, labelsize=8)

top_rev = product_sales.nlargest(10, "revenue").sort_values("revenue")
axes[0, 1].barh(top_rev["Product Name"], top_rev["revenue"], color="#64CCC5")
axes[0, 1].set_title("Top 10 Products by Net Revenue", fontsize=12, fontweight="bold", pad=10)
axes[0, 1].set_xlabel("Net Revenue")

axes[1, 0].barh(category_sales.sort_values("revenue")["Category"], category_sales.sort_values("revenue")["revenue"], color="#053B50")
axes[1, 0].set_title("Net Revenue by Category", fontsize=12, fontweight="bold", pad=10)
axes[1, 0].set_xlabel("Net Revenue")

axes[1, 1].bar(location_sales["Store Location"], location_sales["revenue"], color="#DAA520")
axes[1, 1].set_title("Net Revenue by Store Location", fontsize=12, fontweight="bold", pad=10)
axes[1, 1].set_xlabel("Store Location")
axes[1, 1].set_ylabel("Net Revenue")

fig.tight_layout()
fig.savefig(OUT / "sales_dashboard.png", dpi=160, bbox_inches="tight")
fig.savefig(OUT / "sales_dashboard.pdf", bbox_inches="tight")
plt.close(fig)

print("KPIs:\n", pd.Series(kpis).to_string())
