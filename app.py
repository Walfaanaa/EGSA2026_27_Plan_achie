import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import os
import base64

if os.path.exists("EGSA.png"):
    with open("EGSA.png", "rb") as image_file:
        logo = base64.b64encode(image_file.read()).decode()

    st.markdown(
        f"""
        <div style="text-align:center;">
            <img src="data:image/png;base64,{logo}" width="220">
            <h1 style="color:green; font-size:42px; margin-top:15px; margin-bottom:5px;">
                EGSA 2026/27 Management System
            </h1>
            <p style="color:gray; font-size:20px;">
                Planning, Achievement & Performance Dashboard
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <h1 style="text-align:center; color:green; font-size:42px;">
            EGSA 2026/27 Management System
        </h1>
        <p style="text-align:center; color:gray; font-size:20px;">
            Planning, Achievement & Performance Dashboard
        </p>
        """,
        unsafe_allow_html=True,
    )
# =====================================
# LOAD DATA
# =====================================

st.sidebar.header("📂 Data Source")

data_source = st.sidebar.radio(
    "Choose data source:",
    (
        "Use GitHub Excel File",
        "Upload Excel File"
    )
)

if data_source == "Use GitHub Excel File":

    file_path = "EGSA2026_27_Plan_achie.xlsx"

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        st.sidebar.success("✅ GitHub Excel loaded")
    else:
        st.error("❌ EGSA2026_27_Plan_achie.xlsx not found.")
        st.stop()

else:

    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.sidebar.success("✅ Uploaded file loaded")
    else:
        st.warning("Please upload an Excel file.")
        st.stop()
# =============================
# Clean Columns
# =============================
df.columns = (
    df.columns.str.strip()
    .str.lower()
    .str.replace(" ", "_")
)

numeric_cols = [
    "egsa2026_27_first_half_year_plan",
    "egsa2026_27_first_half_year_achievement",
    "egsa2026_27_monthly_payment",
    "egsa2026_27_second_half_year_plan",
    "egsa2026_27_second_half_year_achievement",
    "fee_charge",
    "volentary_saving",
    "benefit_gain",
    "expenditure",
    "end_2026_achievement"
]

for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# =============================
# Calculations
# =============================
df["annual_achievement"] = (
    df["egsa2026_27_first_half_year_achievement"]
    + df["egsa2026_27_second_half_year_achievement"]
)

df["difference"] = (
    df["end_2026_achievement"]
    - df["annual_achievement"]
)

df["rank"] = (
    df["end_2026_achievement"]
    .rank(method="dense", ascending=False)
    .astype(int)
)

# =============================
# Member Table
# =============================
st.subheader("Member Performance")
st.dataframe(df, use_container_width=True)

# =============================
# =============================
# =============================
# Summary Calculations
# =============================

first_plan = df["egsa2026_27_first_half_year_plan"].sum()
first_ach = df["egsa2026_27_first_half_year_achievement"].sum()

second_plan = df["egsa2026_27_second_half_year_plan"].sum()
second_ach = df["egsa2026_27_second_half_year_achievement"].sum()

monthly = df["egsa2026_27_monthly_payment"].sum()

fee = df["fee_charge"].sum()
saving = df["volentary_saving"].sum()
benefit = df["benefit_gain"].sum()
expense = df["expenditure"].sum()

end_total = df["end_2026_achievement"].sum()

annual_total = (
    first_ach + second_ach
)

net_total = (
    end_total
    + fee
    + saving
    + benefit
    - expense
)


# =============================
# Summary Dashboard
# =============================

st.subheader("📊 EGSA 2026/27 Summary Dashboard")


# Row 1
c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "📅 1st Half Plan",
    f"{first_plan:,.0f}"
)

c2.metric(
    "✅ 1st Half Achievement",
    f"{first_ach:,.0f}"
)

c3.metric(
    "📅 2nd Half Plan",
    f"{second_plan:,.0f}"
)

c4.metric(
    "✅ 2nd Half Achievement",
    f"{second_ach:,.0f}"
)


# Row 2
c5, c6, c7, c8 = st.columns(4)

c5.metric(
    "💵 Monthly Payment",
    f"{monthly:,.0f}"
)

c6.metric(
    "💰 Fee Charge",
    f"{fee:,.0f}"
)

c7.metric(
    "🏦 Voluntary Saving",
    f"{saving:,.0f}"
)

c8.metric(
    "🎁 Benefit Gain",
    f"{benefit:,.0f}"
)


# Row 3
c9, c10, c11 = st.columns(3)


c9.metric(
    "💸 Expenditure",
    f"-{expense:,.0f}"
)


c10.metric(
    "🏆 End 2026 Achievement",
    f"{end_total:,.0f}",
    delta=f"{end_total-annual_total:,.0f}"
)


c11.metric(
    "💎 Net Capital",
    f"{net_total:,.0f}"
)
# =============================
# =============================
# Download Final Report
# =============================

buffer = BytesIO()

with pd.ExcelWriter(
    buffer,
    engine="xlsxwriter"
) as writer:

    final_df.to_excel(
        writer,
        sheet_name="EGSA2026_27_Report",
        index=False
    )

buffer.seek(0)


st.download_button(
    label="📥 Download EGSA 2026/27 Report",
    data=buffer,
    file_name="EGSA2026_27_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
