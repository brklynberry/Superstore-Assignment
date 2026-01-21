import pandas as pd
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

#---Data CLeaning and Processing

xls = pd.ExcelFile("Superstore_Enriched.xlsx")
#print(xls.sheet_names)

df = pd.read_excel("Superstore_Enriched.xlsx", sheet_name="Sample - Superstore")
#print(df.columns)

df.info() #shows us there's no null values in any of the columns
df.isna().sum() #counts number of nulls present in the dataframe
#output for this project was no nulls and no duplicates

df["Order Date"] =  pd.to_datetime(df["Order Date"]) #necessary for date selection visualization
df["Year"] = df["Order Date"].dt.year
#print(df.columns) #useful to quickly see the different columns

#sales_by_region = df.groupby("Region")["Sales"].sum()



#---Streamlit Visualization
#Dashboard Set Up
st.set_page_config(page_title="Superstore Data Visualization", page_icon=None, layout="wide")
st.title(":open_file_folder: Superstore Analysis")
filtered_df = df.copy()

#Date selection option
min_year = df["Order Date"].dt.year.min()
max_year = df["Order Date"].dt.year.max()
year_range = st.slider("Select Year Range:", min_value=int(min_year), max_value=int(max_year), value=(int(min_year), int(max_year)), step=1)
filtered_df = filtered_df[(filtered_df["Order Date"].dt.year >= year_range[0]) & (filtered_df["Order Date"].dt.year <= year_range[1])]


#Region/State/City selection to see respective data
#defaults to entire df if no filters are selected
col3, col4, col5 = st.columns(3) 
with col3:
    region = st.multiselect("Region: ", sorted(df["Region"].unique()))
if not region:
    df2 = df.copy()
else: df2 = df[df["Region"].isin(region)]

with col4:
    state = st.multiselect("State: ", sorted(df2["State"].unique()))
if not state:
    df3 = df2.copy()
else:
    df3 = df2[df2["State"].isin(state)]

with col5:
    city = st.multiselect("City: ", sorted(df3["City"].unique()))
if not city:
    df4 = df3.copy()
else:
    df4 = df3[df3["City"].isin(city)]


if region:
    filtered_df = filtered_df[filtered_df["Region"].isin(region)]
if state:
    filtered_df = filtered_df[filtered_df["State"].isin(state)]
if city:
    filtered_df = filtered_df[filtered_df["City"].isin(city)]


#Statistics Overview - Dependent and changes based on the Region/State/City selection
#Displays total sales, total profit, # of orders, and the profit margin
col1, col2, col3, col4 = st.columns(4) #separates section into 4 columns for each stat

col1.metric("Total Sales", f"${filtered_df['Sales'].sum():,.2f}")
col3.metric("Orders", filtered_df.shape[0]) 
col2.metric("Total Profit", f"${filtered_df['Profit'].sum():,.2f}")
#col4.metric("Profit Margin", f"{(filtered_df["Profit_Margin"].mean() * 100):.2f}%")
col4.metric("Overall Margin", f"{(filtered_df['Profit'].sum() / filtered_df['Sales'].sum() * 100):.2f}%")
 
#output df based on all filters (year range, location)
st.write("Total filtered rows:", filtered_df.shape[0], " of ", df.shape[0], " total.")
st.dataframe(filtered_df)


#Bar chart to visualize categories, donut chart to visualize region
charts_df = filtered_df.groupby(by = ["Category"], as_index = False)["Sales"].sum()
c1, c2 = st.columns(2)
with c1:
    st.subheader("Sales per Category")
    fig = px.bar(charts_df, x = "Category", y = "Sales", text = ['${:,.2f}'.format(x) for x in charts_df["Sales"]],color="Category", template = "seaborn")
    fig.update_layout(showlegend=False)
    st.plotly_chart(fig,use_container_width=True, height = 200)
    st.caption("Total sales per category. Each bar represents the sum of sales for the respective category over the selected filters.")


with c2:
    st.subheader("Sales per Region")
    fig = px.pie(filtered_df, values = "Sales", names = "Region", hole = 0.5)
    fig.update_traces(text = filtered_df["Region"], textposition = "outside")
    st.plotly_chart(fig,use_container_width=True)
    st.caption("Ratio of sales per region. Each slice represents the percentage of total sales made up by each region.")
    

#plot time series sales vs profit
sales_per_month = (filtered_df.set_index("Order Date").resample("M")["Sales"].sum())
profit_per_month = (filtered_df.set_index("Order Date").resample("M")["Profit"].sum())
month_df = pd.DataFrame({"Sales": sales_per_month, "Profit": profit_per_month})

st.subheader("Sales vs. Profit Trends")
st.line_chart(month_df)
st.caption("Relationship between sales and profit by the month.")
#st.line_chart(profit_per_month)

