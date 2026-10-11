from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Sales Analysis Dashboard", layout="wide", initial_sidebar_state="expanded")

HERE = Path(__file__).resolve().parent
CLEAN_FILE = HERE / "output" / "clean_sales_data.csv"
PRODUCTS_FILE = HERE / "output" / "product_summary.csv"

df = pd.read_csv(CLEAN_FILE)
df["Purchase Date"] = pd.to_datetime(df["Purchase Date"])

products_df = pd.read_csv(PRODUCTS_FILE)
category_map = dict(zip(products_df["Product ID"], products_df["Category"]))
prod_name_map = dict(zip(products_df["Product ID"], products_df["Product Name"]))
df["Category"] = df["Product ID"].map(category_map)
df["Product Name"] = df["Product ID"].map(prod_name_map)

st.sidebar.header("Filter Options")

min_date = df["Purchase Date"].min().date()
max_date = df["Purchase Date"].max().date()
date_range = st.sidebar.date_input("Date Range", [min_date, max_date], min_value=min_date, max_value=max_date)

all_locations = sorted(df["Store Location"].dropna().unique().tolist())
selected_locations = st.sidebar.multiselect("Store Location", all_locations, default=all_locations)

all_categories = sorted(df["Category"].dropna().unique().tolist())
selected_categories = st.sidebar.multiselect("Product Category", all_categories, default=all_categories)

all_cust_types = sorted(df["Customer Type"].dropna().unique().tolist())
selected_cust_types = st.sidebar.multiselect("Customer Type", all_cust_types, default=all_cust_types)

filtered = df.copy()
if len(date_range) == 2:
    start_d, end_d = date_range
    filtered = filtered[(filtered["Purchase Date"].dt.date >= start_d) & (filtered["Purchase Date"].dt.date <= end_d)]
if selected_locations:
    filtered = filtered[filtered["Store Location"].isin(selected_locations)]
if selected_categories:
    filtered = filtered[filtered["Category"].isin(selected_categories)]
if selected_cust_types:
    filtered = filtered[filtered["Customer Type"].isin(selected_cust_types)]

st.title("Executive Sales Analysis Dashboard")
st.caption(f"Analyzing {len(filtered):,} transactions  |  Source: Sales_Data.xlsx")

gross_sales = filtered["Total Price"].sum()
discounts = filtered["Discount Amount"].sum()
net_revenue = filtered["Final Price"].sum()
units_sold = filtered["Quantity"].sum()
orders = filtered["Order ID"].nunique()

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Gross Sales", f"{gross_sales:,.2f}")
col2.metric("Total Discounts", f"{discounts:,.2f}")
col3.metric("Net Revenue", f"{net_revenue:,.2f}")
col4.metric("Units Sold", f"{units_sold:,}")
col5.metric("Total Orders", f"{orders:,}")

prod_grp = filtered.groupby("Product Name", as_index=False).agg(
    Revenue=("Final Price", "sum"),
    Units=("Quantity", "sum")
)
top_rev_prod = prod_grp.sort_values("Revenue", ascending=False).iloc[0]["Product Name"] if len(prod_grp) > 0 else "N/A"
top_rev_val = prod_grp.sort_values("Revenue", ascending=False).iloc[0]["Revenue"] if len(prod_grp) > 0 else 0
top_unit_prod = prod_grp.sort_values("Units", ascending=False).iloc[0]["Product Name"] if len(prod_grp) > 0 else "N/A"
top_unit_val = prod_grp.sort_values("Units", ascending=False).iloc[0]["Units"] if len(prod_grp) > 0 else 0

st.info(f"Top Product by Revenue: **{top_rev_prod}** ({top_rev_val:,.2f})  |  Top Product by Volume: **{top_unit_prod}** ({top_unit_val:,} units)")

st.markdown("---")

c_left, c_right = st.columns(2)

monthly = filtered.assign(Month=filtered["Purchase Date"].dt.to_period("M").astype(str)).groupby("Month", as_index=False).agg(
    Revenue=("Final Price", "sum"),
    Orders=("Order ID", "nunique"),
)
fig_monthly = px.bar(
    monthly,
    x="Month",
    y="Revenue",
    title="Monthly Net Revenue Trend",
    labels={"Revenue": "Net Revenue", "Month": "Month"},
    color_discrete_sequence=["#176B87"],
)
fig_monthly.update_layout(
    margin=dict(l=20, r=20, t=50, b=40),
    height=380,
    xaxis_tickangle=-45,
    yaxis=dict(tickformat=",.0f")
)
c_left.plotly_chart(fig_monthly, use_container_width=True)

cat_perf = filtered.groupby("Category", as_index=False).agg(Revenue=("Final Price", "sum")).sort_values("Revenue", ascending=True)
fig_cat = px.bar(
    cat_perf,
    x="Revenue",
    y="Category",
    orientation="h",
    title="Net Revenue by Category",
    labels={"Revenue": "Net Revenue", "Category": "Category"},
    color_discrete_sequence=["#64CCC5"],
)
fig_cat.update_layout(
    margin=dict(l=20, r=20, t=50, b=20),
    height=380,
    xaxis=dict(tickformat=",.0f")
)
c_right.plotly_chart(fig_cat, use_container_width=True)

c_prod, c_store = st.columns(2)

prod_perf = filtered.groupby("Product Name", as_index=False).agg(Revenue=("Final Price", "sum")).sort_values("Revenue", ascending=True).tail(10)
fig_prod = px.bar(
    prod_perf,
    x="Revenue",
    y="Product Name",
    orientation="h",
    title="Top 10 Products by Net Revenue",
    labels={"Revenue": "Net Revenue", "Product Name": "Product"},
    color_discrete_sequence=["#053B50"],
)
fig_prod.update_layout(
    margin=dict(l=20, r=20, t=50, b=20),
    height=380,
    xaxis=dict(tickformat=",.0f")
)
c_prod.plotly_chart(fig_prod, use_container_width=True)

store_perf = filtered.groupby("Store Location", as_index=False).agg(Revenue=("Final Price", "sum")).sort_values("Revenue", ascending=False)
fig_store = px.bar(
    store_perf,
    x="Store Location",
    y="Revenue",
    title="Net Revenue by Store Location",
    labels={"Revenue": "Net Revenue", "Store Location": "Store Location"},
    color_discrete_sequence=["#DAA520"],
)
fig_store.update_layout(
    margin=dict(l=20, r=20, t=50, b=20),
    height=380,
    yaxis=dict(tickformat=",.0f")
)
c_store.plotly_chart(fig_store, use_container_width=True)

with st.expander("Explore Filtered Data Records"):
    st.dataframe(filtered, use_container_width=True)
