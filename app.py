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
# =============================
# Upload Excel
# =============================
uploaded = st.file_uploader(
    "Upload EGSA Excel File",
    type=["xlsx", "xls"]
)

if uploaded is None:
    st.stop()

df = pd.read_excel(uploaded)

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
# Summary
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

annual_total = first_ach + second_ach

net_total = (
    end_total
    + fee
    + saving
    + benefit
    - expense
)

st.subheader("Summary")

c1,c2,c3,c4,c5 = st.columns(5)

c1.metric("First Half Plan",f"{first_plan:,.0f}")
c2.metric("Second Half Plan",f"{second_plan:,.0f}")
c3.metric("Annual Achievement",f"{annual_total:,.0f}")
c4.metric(
    "End 2026 Achievement",
    f"{end_total:,.0f}",
    delta=f"{end_total-annual_total:,.0f}"
)
c5.metric("Net Capital",f"{net_total:,.0f}")

# =============================
# Comparison Chart
# =============================
st.subheader("Achievement Comparison")

fig, ax = plt.subplots(figsize=(7,5))

labels = [
    "Annual Achievement",
    "End 2026 Achievement"
]

values = [
    annual_total,
    end_total
]

ax.bar(labels, values)

for i,v in enumerate(values):
    ax.text(i,v,f"{v:,.0f}",ha="center")

st.pyplot(fig)

# =============================
# Member Ranking
# =============================
st.subheader("Top Members")

top = (
    df.sort_values(
        "end_2026_achievement",
        ascending=False
    )
    [["id","name","end_2026_achievement","rank"]]
)

st.dataframe(top,use_container_width=True)

# =============================
# Final Report
# =============================
total = pd.DataFrame({
    "name":["TOTAL"],
    "egsa2026_27_first_half_year_plan":[first_plan],
    "egsa2026_27_first_half_year_achievement":[first_ach],
    "egsa2026_27_second_half_year_plan":[second_plan],
    "egsa2026_27_second_half_year_achievement":[second_ach],
    "egsa2026_27_monthly_payment":[monthly],
    "fee_charge":[fee],
    "volentary_saving":[saving],
    "benefit_gain":[benefit],
    "expenditure":[expense],
    "end_2026_achievement":[end_total],
    "difference":[end_total-annual_total]
})

final_df = pd.concat([df,total],ignore_index=True)

st.subheader("Final Report")
st.dataframe(final_df,use_container_width=True)

# =============================
# Download
# =============================
buffer = BytesIO()

with pd.ExcelWriter(buffer,engine="openpyxl") as writer:
    final_df.to_excel(writer,index=False)

buffer.seek(0)

st.download_button(
    "Download Report",
    data=buffer,
    file_name="EGSA2026_27_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
