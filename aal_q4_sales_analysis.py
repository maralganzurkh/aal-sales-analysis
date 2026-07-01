#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AAL Q4-2020 Sales Analysis
---------------------------------
- Loads AusApparalSales4thQrt2020.csv
- Cleans and inspects data
- Creates descriptive statistics
- Aggregates sales by state and group
- Builds daily/weekly/monthly/quarterly series
- Analyses time-of-day categories
- Exports an Excel report with all key tables
- Plots charts (matplotlib only)
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

# ---------- config ----------
# Put the CSV next to this script, OR update csv_path below.
DEFAULT_NAME = "AusApparalSales4thQrt2020.csv"
here = Path(".").resolve()
csv_candidates = [
    here / DEFAULT_NAME,
    Path("/mnt/data") / DEFAULT_NAME  # fallback for hosted environments
]
csv_path = None
for c in csv_candidates:
    if c.exists():
        csv_path = c
        break
if csv_path is None:
    raise FileNotFoundError(
        f"Could not find {DEFAULT_NAME} in {here} or /mnt/data. "
        "Place the CSV next to this script or update csv_path."
    )

# Output Excel
output_path = here / "AAL_Q4_2020_Sales_Report.xlsx"

# ---------- load data ----------
df = pd.read_csv(csv_path)

# Keep a copy of the original metadata
orig_shape = df.shape
orig_cols = df.columns.tolist()
orig_dtypes = df.dtypes.astype(str).to_dict()

# Standardise column names
df.columns = [c.strip().replace(" ", "_") for c in df.columns]

# Expecting columns like: Date, Time, State, Group, Unit, Sales
# Attempt to build a DateTime from Date/Time
date_col = None
if "Date" in df.columns:
    date_col = "Date"
elif "date" in df.columns:
    date_col = "date"

time_col = None
if "Time" in df.columns:
    time_col = "Time"
elif "time" in df.columns:
    time_col = "time"

if date_col and time_col:
    df["DateTime"] = pd.to_datetime(df[date_col].astype(str) + " " + df[time_col].astype(str), errors="coerce", dayfirst=True)
elif date_col:
    df["DateTime"] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
else:
    df["DateTime"] = pd.NaT

# Extract calendar fields (will be NA if DateTime missing)
df["Date_only"] = df["DateTime"].dt.date
df["Week"] = df["DateTime"].dt.isocalendar().week.astype("Int64")
df["Month"] = df["DateTime"].dt.month.astype("Int64")
df["Quarter"] = df["DateTime"].dt.quarter.astype("Int64")

# Identify key columns
state_col = "State" if "State" in df.columns else None
group_col = "Group" if "Group" in df.columns else None
sales_col = "Sales" if "Sales" in df.columns else None
units_col = "Unit" if "Unit" in df.columns else None

# ---------- basic data quality ----------
missing_summary = df.isna().sum().to_frame("missing_count")
missing_summary["missing_pct"] = (missing_summary["missing_count"] / len(df) * 100).round(2)

incorrect_flags = {}
if sales_col:
    incorrect_flags["negative_sales_count"] = int((df[sales_col] < 0).sum())
if units_col:
    incorrect_flags["negative_units_count"] = int((df[units_col] < 0).sum())

# ---------- descriptive stats ----------
desc_stats = {}
if sales_col:
    desc_stats["Sales"] = {
        "count": int(df[sales_col].count()),
        "mean": float(df[sales_col].mean()),
        "median": float(df[sales_col].median()),
        "mode": float(df[sales_col].mode().iloc[0]) if not df[sales_col].mode().empty else np.nan,
        "std": float(df[sales_col].std(ddof=1)) if df[sales_col].count() > 1 else np.nan,
        "min": float(df[sales_col].min()),
        "max": float(df[sales_col].max()),
    }
if units_col:
    desc_stats["Units"] = {
        "count": int(df[units_col].count()),
        "mean": float(df[units_col].mean()),
        "median": float(df[units_col].median()),
        "mode": float(df[units_col].mode().iloc[0]) if not df[units_col].mode().empty else np.nan,
        "std": float(df[units_col].std(ddof=1)) if df[units_col].count() > 1 else np.nan,
        "min": float(df[units_col].min()),
        "max": float(df[units_col].max()),
    }
desc_df = pd.DataFrame(desc_stats).round(2)

# ---------- normalisation for modelling ----------
norm_df = df.copy()
for col in [c for c in [sales_col, units_col] if c]:
    cmin, cmax = df[col].min(), df[col].max()
    if pd.notna(cmin) and pd.notna(cmax) and cmax != cmin:
        norm_df[col + "_norm"] = (df[col] - cmin) / (cmax - cmin)
    else:
        norm_df[col + "_norm"] = np.nan

# ---------- aggregations ----------
state_totals = None
group_totals = None
daily = weekly = monthly = quarterly = None
time_of_day_totals = None

if sales_col and state_col:
    state_totals = (
        df.groupby(state_col, dropna=False)[sales_col]
          .sum()
          .sort_values(ascending=False)
          .to_frame("Total_Sales")
    )

if sales_col and group_col:
    group_totals = (
        df.groupby(group_col, dropna=False)[sales_col]
          .sum()
          .sort_values(ascending=False)
          .to_frame("Total_Sales")
    )

if sales_col:
    if "Date_only" in df.columns:
        daily = df.groupby("Date_only")[sales_col].sum().reset_index().sort_values("Date_only")
    if "Week" in df.columns and df["Week"].notna().any():
        weekly = df.groupby("Week")[sales_col].sum().reset_index().sort_values("Week")
    if "Month" in df.columns and df["Month"].notna().any():
        monthly = df.groupby("Month")[sales_col].sum().reset_index().sort_values("Month")
    if "Quarter" in df.columns and df["Quarter"].notna().any():
        quarterly = df.groupby("Quarter")[sales_col].sum().reset_index().sort_values("Quarter")

# Many versions of the dataset label time categorically (e.g., Morning/Afternoon/Evening)
if "Time" in df.columns:
    time_of_day_totals = df.groupby("Time")[sales_col].sum().reset_index().sort_values("Sales", ascending=False)

# ---------- export Excel ----------
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Raw Data", index=False)
    cols_norm = [c for c in df.columns] + [c + "_norm" for c in [x for x in ["Sales","Unit"] if x in df.columns]]
    norm_df[cols_norm].to_excel(writer, sheet_name="Normalized", index=False)
    desc_df.to_excel(writer, sheet_name="Descriptive Stats")
    if state_totals is not None:
        state_totals.to_excel(writer, sheet_name="State Totals")
    if group_totals is not None:
        group_totals.to_excel(writer, sheet_name="Group Totals")
    if daily is not None:
        daily.to_excel(writer, sheet_name="Daily Totals", index=False)
    if weekly is not None:
        weekly.to_excel(writer, sheet_name="Weekly Totals", index=False)
    if monthly is not None:
        monthly.to_excel(writer, sheet_name="Monthly Totals", index=False)
    if quarterly is not None:
        quarterly.to_excel(writer, sheet_name="Quarterly Totals", index=False)
    if time_of_day_totals is not None:
        time_of_day_totals.to_excel(writer, sheet_name="Time-of-Day Totals", index=False)

print(f"Excel written to: {output_path.resolve()}")

# ---------- quick leaderboard & insights ----------
if state_totals is not None and not state_totals.empty:
    sr = state_totals.reset_index().rename(columns={state_col:"State"})
    sr["State"] = sr["State"].str.strip()
    sr = sr.sort_values("Total_Sales", ascending=False)

    top3 = sr.head(3)
    bottom3 = sr.tail(3)

    print("\nTop 3 states by sales (Q4 2020):")
    print(top3.to_string(index=False))

    print("\nBottom 3 states by sales (Q4 2020):")
    print(bottom3.to_string(index=False))

    # Strongest/weakest group per bottom state
    if group_col:
        per_state_group = (
            df.groupby([state_col, group_col])[sales_col]
              .sum()
              .reset_index()
        )
        per_state_group[state_col] = per_state_group[state_col].str.strip()
        per_state_group[group_col] = per_state_group[group_col].str.strip()

        for st in bottom3["State"].tolist():
            sub = per_state_group[per_state_group[state_col]==st]
            if not sub.empty:
                best = sub.sort_values(sales_col, ascending=False).iloc[0]
                worst = sub.sort_values(sales_col, ascending=True).iloc[0]
                print(f"\n{st}: best group = {best[group_col]} ({int(best[sales_col]):,}), "
                      f"weakest group = {worst[group_col]} ({int(worst[sales_col]):,})")

# ---------- plots (one figure per chart; no seaborn; no custom colours) ----------
plt.close("all")

# State × Group stacked bar
if sales_col and state_col and group_col:
    pivot_sg = df.pivot_table(index=state_col, columns=group_col, values=sales_col, aggfunc="sum").fillna(0)
    ax = pivot_sg.plot(kind="bar", stacked=True, figsize=(10,6))
    ax.set_title("State-wise Sales by Group (Q4 2020)")
    ax.set_xlabel("State")
    ax.set_ylabel("Total Sales")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.show()

# Group × State grouped bar
if sales_col and state_col and group_col:
    pivot_gs = df.pivot_table(index=group_col, columns=state_col, values=sales_col, aggfunc="sum").fillna(0)
    ax = pivot_gs.plot(kind="bar", figsize=(10,6))
    ax.set_title("Group-wise Sales Across States (Q4 2020)")
    ax.set_xlabel("Group")
    ax.set_ylabel("Total Sales")
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.show()

# Time-of-day bar (if available)
if time_of_day_totals is not None and not time_of_day_totals.empty:
    ax = time_of_day_totals.plot(x="Time", y="Sales", kind="bar", figsize=(8,5))
    ax.set_title("Sales by Time-of-Day (Q4 2020)")
    ax.set_xlabel("Time of Day")
    ax.set_ylabel("Total Sales")
    plt.tight_layout()
    plt.show()

# Daily line
if daily is not None and not daily.empty:
    ax = daily.plot(x="Date_only", y=sales_col, kind="line", figsize=(10,5))
    ax.set_title("Daily Sales (Q4 2020)")
    ax.set_xlabel("Date")
    ax.set_ylabel("Total Sales")
    plt.tight_layout()
    plt.show()

# Weekly line
if weekly is not None and not weekly.empty:
    ax = weekly.plot(x="Week", y=sales_col, kind="line", figsize=(10,5))
    ax.set_title("Weekly Sales (Q4 2020)")
    ax.set_xlabel("ISO Week Number")
    ax.set_ylabel("Total Sales")
    plt.tight_layout()
    plt.show()

# Monthly bar
if monthly is not None and not monthly.empty:
    ax = monthly.plot(x="Month", y=sales_col, kind="bar", figsize=(8,5))
    ax.set_title("Monthly Sales (Q4 2020)")
    ax.set_xlabel("Month")
    ax.set_ylabel("Total Sales")
    plt.tight_layout()
    plt.show()

# Quarterly bar
if quarterly is not None and not quarterly.empty:
    ax = quarterly.plot(x="Quarter", y=sales_col, kind="bar", figsize=(6,4))
    ax.set_title("Quarterly Sales (Q4 2020)")
    ax.set_xlabel("Quarter")
    ax.set_ylabel("Total Sales")
    plt.tight_layout()
    plt.show()

# Box plots
if sales_col:
    plt.figure(figsize=(6,5))
    plt.boxplot(df[sales_col].dropna())
    plt.title("Box Plot of Sales (Q4 2020)")
    plt.ylabel("Sales")
    plt.tight_layout()
    plt.show()

if units_col:
    plt.figure(figsize=(6,5))
    plt.boxplot(df[units_col].dropna())
    plt.title("Box Plot of Units (Q4 2020)")
    plt.ylabel("Units")
    plt.tight_layout()
    plt.show()
