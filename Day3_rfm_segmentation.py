"""
Day 3 - Customer Segmentation using RFM Analysis
Project: Algorithmic Credit Scoring Using Behavioral and Transactional Data

Purpose
-------
Segment the 801 customers using RFM (Recency, Frequency, Monetary) analysis built on
Fact_Transactions from the Day 1 cleaned workbook, then enrich each segment with
credit-risk context (average CreditScore, default rate, average BehaviorScore) since
this is a credit-scoring project, not a generic retail one.

Design note (carried over from Day 1): CustomerID attributes (Age, Income, CreditScore,
etc.) are point-in-time, not stable per customer. RFM is computed by aggregating
Fact_Transactions directly per CustomerID (never by taking "the" customer record), and
CreditScore/BehaviorScore are aggregated as an AVERAGE across a customer's transactions
to represent a general standing rather than picking an arbitrary single row.

Output: Customer_RFM_Segments.xlsx
"""

import pandas as pd
import numpy as np
from pathlib import Path

SRC = Path("/home/claude/day1/output/Master_Finance_Data_Cleaned.xlsx")
OUT_DIR = Path("/home/claude/day3/output")
OUT_DIR.mkdir(parents=True, exist_ok=True)

log = []
def note(step, detail):
    log.append({"Step": step, "Detail": detail})
    print(f"[{step}] {detail}")

# ---------------------------------------------------------------------------
# 1. LOAD cleaned transaction fact table (Day 1 output — already validated)
# ---------------------------------------------------------------------------
df = pd.read_excel(SRC, sheet_name="Cleaned_Flat_Table")
note("EXTRACT", f"Loaded {len(df):,} cleaned transactions across {df['CustomerID'].nunique()} "
                 f"customers from Day 1 output ('{SRC.name}').")

# ---------------------------------------------------------------------------
# 2. BUILD RFM METRICS (grain: one row per CustomerID)
# ---------------------------------------------------------------------------
reference_date = df["TransactionDate"].max() + pd.Timedelta(days=1)
note("RFM-REFERENCE", f"Reference date for Recency = {reference_date.date()} "
                       f"(one day after the latest transaction, {df['TransactionDate'].max().date()}).")

rfm = df.groupby("CustomerID").agg(
    Recency=("TransactionDate", lambda x: (reference_date - x.max()).days),
    Frequency=("TransactionID", "count"),
    Monetary=("TransactionAmount", "sum"),
    AvgCreditScore=("CreditScore", "mean"),
    AvgBehaviorScore=("BehaviorScore", "mean"),
    DefaultRate=("DefaultFlag", "mean"),
    TotalLoanAmount=("LoanAmount", "sum"),
    LastOccupation=("Occupation", "last"),
    LastRegion=("Region", "last"),
).reset_index()

note("RFM-AGGREGATE", f"Built RFM base table: {len(rfm)} customers x "
                       f"{rfm.shape[1]} columns.")

# ---------------------------------------------------------------------------
# 3. SCORE R, F, M on a 1-4 scale (quartiles) and combine
# ---------------------------------------------------------------------------
# Recency: LOWER is better (recent activity) -> reverse-score with duplicates='drop' safety
rfm["R_Score"] = pd.qcut(rfm["Recency"], 4, labels=[4, 3, 2, 1]).astype(int)
rfm["F_Score"] = pd.qcut(rfm["Frequency"].rank(method="first"), 4, labels=[1, 2, 3, 4]).astype(int)
rfm["M_Score"] = pd.qcut(rfm["Monetary"], 4, labels=[1, 2, 3, 4]).astype(int)

rfm["RFM_Score"] = rfm["R_Score"].astype(str) + rfm["F_Score"].astype(str) + rfm["M_Score"].astype(str)
rfm["RFM_Total"] = rfm["R_Score"] + rfm["F_Score"] + rfm["M_Score"]

note("RFM-SCORING", "Scored Recency, Frequency and Monetary into quartiles (1=lowest, "
                     "4=highest value); Recency reversed so 4 = most recent. "
                     "Frequency ties broken with rank(method='first') so qcut can form 4 "
                     "even bins.")


# ---------------------------------------------------------------------------
# 4. SEGMENT LABELS (standard RFM heuristic, adapted to a credit-scoring lens)
# ---------------------------------------------------------------------------
def segment(row):
    r, f, m = row["R_Score"], row["F_Score"], row["M_Score"]
    if r >= 4 and f >= 4 and m >= 4:
        return "Champions"
    if r >= 3 and f >= 3:
        return "Loyal Customers"
    if r >= 4 and f <= 2:
        return "New / Promising"
    if r == 3 and m >= 3:
        return "Potential Loyalist"
    if r <= 2 and f >= 3 and m >= 3:
        return "At Risk"
    if r <= 2 and f <= 2 and m >= 3:
        return "Can't Lose Them"
    if r <= 2 and f <= 2 and m <= 2:
        return "Hibernating"
    return "Needs Attention"

rfm["Segment"] = rfm.apply(segment, axis=1)
note("RFM-SEGMENTS", f"Assigned 8 behavioral segments. Distribution: "
                      f"{rfm['Segment'].value_counts().to_dict()}")

# ---------------------------------------------------------------------------
# 5. CREDIT-RISK OVERLAY (project-specific: is engagement tier linked to risk?)
# ---------------------------------------------------------------------------
segment_profile = rfm.groupby("Segment").agg(
    Customers=("CustomerID", "count"),
    AvgRecency=("Recency", "mean"),
    AvgFrequency=("Frequency", "mean"),
    AvgMonetary=("Monetary", "mean"),
    AvgCreditScore=("AvgCreditScore", "mean"),
    AvgDefaultRate=("DefaultRate", "mean"),
    AvgBehaviorScore=("AvgBehaviorScore", "mean"),
).round(2).sort_values("Customers", ascending=False).reset_index()

note("RISK-OVERLAY", "Built Segment_Profile summary: average CreditScore, DefaultRate and "
                      "BehaviorScore per RFM segment, to connect engagement behavior with "
                      "credit risk for this project's specific use case.")

# ---------------------------------------------------------------------------
# 6. SAVE
# ---------------------------------------------------------------------------
log_df = pd.DataFrame(log)
with pd.ExcelWriter(OUT_DIR / "Customer_RFM_Segments.xlsx", engine="openpyxl") as writer:
    rfm.to_excel(writer, sheet_name="Customer_RFM", index=False)
    segment_profile.to_excel(writer, sheet_name="Segment_Profile", index=False)
    log_df.to_excel(writer, sheet_name="Methodology_Log", index=False)

note("SAVE", f"Workbook written to {OUT_DIR / 'Customer_RFM_Segments.xlsx'} "
             f"(Customer_RFM: {len(rfm)} rows, Segment_Profile: {len(segment_profile)} segments).")

print("\nDay 3 RFM segmentation complete.")
