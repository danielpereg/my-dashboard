#!/usr/bin/env python3
"""
app.py

Streamlit app to visualize `Client_Dashboard.xlsx`.

Features:
- Wide layout with a polished header and premium styling
- Top KPIs: Total Revenue, Total Clients, Average Purchase
- Interactive Plotly bar chart of `purchase_amount` per `customer_name`
- Sidebar multiselect filter for `customer_name`
- Raw data displayed below using `st.dataframe`
- Error handling for missing or invalid files

Run:
    streamlit run app.py
"""
from __future__ import annotations

import logging
import os
from typing import Optional

import pandas as pd
import plotly.express as px
import streamlit as st


# ---------------------- Configuration ----------------------
DATA_FILE = "Client_Dashboard.xlsx"

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


# ---------------------- Utility functions ----------------------

def load_data(path: str) -> Optional[pd.DataFrame]:
    """Load Excel data into a DataFrame. Returns None if file missing or unreadable.

    Uses `openpyxl` engine for compatibility with xlsxwriter outputs.
    """
    if not os.path.exists(path):
        logger.error("Data file not found: %s", path)
        return None

    try:
        df = pd.read_excel(path, engine="openpyxl")
        logger.info("Loaded %d rows and %d columns from %s", len(df), len(df.columns), path)
        return df
    except Exception as exc:
        logger.exception("Failed to read Excel file: %s", exc)
        return None


def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure required columns exist and coerce types as needed.

    - If `purchase_amount` doesn't exist, create it with zeros.
    - If `customer_name` doesn't exist, create a fallback column.
    """
    df = df.copy()

    if "purchase_amount" not in df.columns:
        logger.warning("'purchase_amount' column not found; creating with zeros")
        df["purchase_amount"] = 0.0
    else:
        df["purchase_amount"] = pd.to_numeric(df["purchase_amount"], errors="coerce").fillna(0.0)

    if "customer_name" not in df.columns:
        logger.warning("'customer_name' column not found; creating fallback names")
        df["customer_name"] = [f"Client {i+1}" for i in range(len(df))]
    else:
        df["customer_name"] = df["customer_name"].astype(str).fillna("")

    return df


# ---------------------- Streamlit App ----------------------

st.set_page_config(page_title="Client Performance Analytics", layout="wide", page_icon="📊")

# Custom page style for a luxury look
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(180deg, #eef3fb 0%, #f9fbff 100%);
        color: #0f172a;
    }
    .css-1d391kg {
        padding-top: 1rem;
    }
    div[data-testid="stMetric"] {
        background: rgba(47, 85, 151, 0.12);
        border-radius: 26px;
        padding: 1rem 1.2rem;
        box-shadow: 0 18px 40px rgba(15, 23, 42, 0.08);
        border: 1px solid rgba(47, 85, 151, 0.14);
    }
    .css-1avcm0n {
        background: rgba(255, 255, 255, 0.85);
        box-shadow: 0 18px 40px rgba(15, 23, 42, 0.08);
        border-radius: 28px;
        border: 1px solid rgba(15, 23, 42, 0.08);
        padding: 1.5rem;
    }
    .stDataFrame > div {
        border-radius: 24px;
        overflow: hidden;
    }
    .stDataFrame table {
        border-collapse: separate;
        border-spacing: 0px 8px;
    }
    .stDataFrame th {
        background: rgba(47, 85, 151, 0.08) !important;
        color: #12284b !important;
    }
    .stDataFrame td {
        background: rgba(255, 255, 255, 0.95) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header
header_col, badge_col = st.columns([3, 1], gap="large")
with header_col:
    st.markdown(
        """
        <div style='padding: 0.75rem 0;'>
            <h1 style='margin:0; color:#12284b; letter-spacing: -0.05em;'>Client Performance Analytics</h1>
            <p style='margin:0.75rem 0 0 0; color:#475569; font-size:1rem; max-width:720px;'>A premium analytics experience that highlights revenue performance, client value, and purchase trends with a modern executive interface.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with badge_col:
    st.markdown(
        """
        <div style='background: linear-gradient(135deg, #2F5597 0%, #4D7EE6 100%); color:white; border-radius:24px; padding:1.25rem; text-align:center; box-shadow:0 24px 60px rgba(47,85,151,0.18);'>
            <div style='font-size:0.85rem; opacity:0.85; letter-spacing: 0.08em;'>LUXURY INSIGHTS</div>
            <div style='font-size:1.7rem; font-weight:700; margin-top:0.75rem;'>Executive Edition</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("---")

# Load data
data = load_data(DATA_FILE)
if data is None:
    st.error(f"Data file not found or unreadable: `{DATA_FILE}`. Please ensure the file exists.")
    st.stop()

# Prepare data
df = ensure_columns(data)

# Sidebar filters
st.sidebar.header("Filters")
st.sidebar.write("Refine the dashboard by customer or insight.")
customers = sorted(df["customer_name"].unique())
selected_customers = st.sidebar.multiselect("Customer name", customers, default=customers)

# Filtered dataframe
df_filtered = df[df["customer_name"].isin(selected_customers)].copy()

# KPI section
total_revenue = df_filtered["purchase_amount"].sum()
unique_clients = df_filtered["customer_name"].nunique()
avg_purchase = df_filtered["purchase_amount"].mean()

total_revenue_fmt = f"${total_revenue:,.2f}"
avg_purchase_fmt = f"${avg_purchase:,.2f}" if not pd.isna(avg_purchase) else "$0.00"

kpi1, kpi2, kpi3 = st.columns([1, 1, 1], gap="large")
with kpi1:
    st.metric(label="Total Revenue", value=total_revenue_fmt, delta=f"{(total_revenue/unique_clients if unique_clients else 0):,.2f} avg")
with kpi2:
    st.metric(label="Total Clients", value=str(unique_clients), delta="Audience scale")
with kpi3:
    st.metric(label="Average Purchase", value=avg_purchase_fmt, delta="Executive benchmark")

# Insights row
insight_1, insight_2, insight_3 = st.columns(3, gap="large")
if not df_filtered.empty:
    top_customer = df_filtered.groupby("customer_name", as_index=False)["purchase_amount"].sum().sort_values("purchase_amount", ascending=False).iloc[0]
    revenue_share = top_customer["purchase_amount"] / total_revenue * 100 if total_revenue else 0
    with insight_1:
        st.markdown(
            f"<div style='background:#FFFFFF; border-radius:24px; padding:1rem; box-shadow:0 18px 40px rgba(15,23,42,0.05);'>"
            f"<div style='color:#475569; font-size:0.9rem; margin-bottom:0.5rem;'>Top Client</div>"
            f"<div style='font-size:1.2rem; font-weight:700; color:#12284b;'>{top_customer['customer_name']}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    with insight_2:
        st.markdown(
            f"<div style='background:#FFFFFF; border-radius:24px; padding:1rem; box-shadow:0 18px 40px rgba(15,23,42,0.05);'>"
            f"<div style='color:#475569; font-size:0.9rem; margin-bottom:0.5rem;'>Top Purchase</div>"
            f"<div style='font-size:1.2rem; font-weight:700; color:#12284b;'>${top_customer['purchase_amount']:,.2f}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    with insight_3:
        st.markdown(
            f"<div style='background:#FFFFFF; border-radius:24px; padding:1rem; box-shadow:0 18px 40px rgba(15,23,42,0.05);'>"
            f"<div style='color:#475569; font-size:0.9rem; margin-bottom:0.5rem;'>Share of Revenue</div>"
            f"<div style='font-size:1.2rem; font-weight:700; color:#12284b;'>{revenue_share:,.1f}%</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
else:
    insight_1.write("No data available for insights.")
    insight_2.write("")
    insight_3.write("")

st.markdown("---")

# Visualization
if df_filtered.empty:
    st.info("No data available for selected filters.")
else:
    chart_df = (
        df_filtered.groupby("customer_name", as_index=False)["purchase_amount"]
        .sum()
        .sort_values("purchase_amount", ascending=False)
    )

    fig = px.bar(
        chart_df,
        x="customer_name",
        y="purchase_amount",
        labels={"customer_name": "Customer", "purchase_amount": "Purchase Amount"},
        title="Purchase Amount by Customer",
        color="purchase_amount",
        color_continuous_scale=["#2F5597", "#7CA8F7"],
        template="plotly_white",
    )
    fig.update_traces(marker_line_width=0, marker_opacity=0.96)
    fig.update_layout(
        plot_bgcolor="rgba(255,255,255,0)",
        paper_bgcolor="rgba(255,255,255,0)",
        title={"x": 0.0, "xanchor": "left", "font": {"size": 22, "color": "#0f172a"}},
        font={"family": "Inter, sans-serif", "color": "#0f172a"},
        margin={"t": 50, "r": 10, "l": 10, "b": 20},
        coloraxis_showscale=False,
    )
    fig.update_xaxes(showgrid=False, tickangle=-45, title_text=None)
    fig.update_yaxes(gridcolor="rgba(47,85,151,0.08)", title_text="Purchase Amount")
    fig.update_layout(height=520)

    st.plotly_chart(fig, use_container_width=True)

# Styled data table
st.subheader("Detailed Client Revenue Table")
st.markdown("Explore the filtered client records and revenue breakdown below.")
if df_filtered.empty:
    st.write("No data to display.")
else:
    st.dataframe(df_filtered, use_container_width=True)

st.caption("Interactive dashboard with client filtering, executive KPIs, and a refined premium layout.")
