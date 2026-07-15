import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

plt.rcParams["font.family"] = "DejaVu Sans"
CHART_DIR = "/home/claude/day3/charts"

rfm = pd.read_excel("/home/claude/day3/output/Customer_RFM_Segments.xlsx", sheet_name="Customer_RFM")
profile = pd.read_excel("/home/claude/day3/output/Customer_RFM_Segments.xlsx", sheet_name="Segment_Profile")

COLORS = ["#2E5266", "#6E8898", "#9FB8AD", "#D7B377", "#BF4E30", "#4C6444", "#8E7CC3", "#C9A0A0"]
seg_order = profile.sort_values("Customers", ascending=False)["Segment"].tolist()
seg_colors = dict(zip(seg_order, COLORS[:len(seg_order)]))

# 1. Segment size distribution
fig, ax = plt.subplots(figsize=(8,4.8))
counts = rfm["Segment"].value_counts().reindex(seg_order)
bars = ax.barh(counts.index[::-1], counts.values[::-1], color=[seg_colors[s] for s in counts.index[::-1]])
for b, v in zip(bars, counts.values[::-1]):
    ax.text(v + 3, b.get_y()+b.get_height()/2, f"{v} ({v/len(rfm):.1%})", va="center", fontsize=9)
ax.set_title("Customer Count by RFM Segment (n=801)", fontsize=13, fontweight="bold")
ax.set_xlabel("Customers")
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/01_segment_sizes.png", dpi=150)
plt.close()

# 2. RFM scatter: Recency vs Monetary, sized by Frequency, colored by segment
fig, ax = plt.subplots(figsize=(8,5.5))
for seg in seg_order:
    sub = rfm[rfm["Segment"]==seg]
    ax.scatter(sub["Recency"], sub["Monetary"], s=sub["Frequency"]*12, alpha=0.6,
               color=seg_colors[seg], label=seg, edgecolor="white", linewidth=0.4)
ax.set_xlabel("Recency (days since last transaction) — lower is more recent")
ax.set_ylabel("Monetary (total spend, ₹)")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x/1e6:.1f}M" if x>=1e6 else f"{x/1e3:.0f}K"))
ax.set_title("Customer Landscape: Recency vs Monetary\n(bubble size = Frequency)", fontsize=13, fontweight="bold")
ax.legend(loc="upper right", fontsize=8, framealpha=0.9)
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/02_rfm_scatter.png", dpi=150)
plt.close()

# 3. Default rate by segment (credit risk overlay)
fig, ax = plt.subplots(figsize=(8,4.8))
p = profile.sort_values("AvgDefaultRate", ascending=False)
bars = ax.bar(p["Segment"], p["AvgDefaultRate"]*100, color=[seg_colors[s] for s in p["Segment"]])
for b, v in zip(bars, p["AvgDefaultRate"]):
    ax.text(b.get_x()+b.get_width()/2, v*100+0.4, f"{v:.1%}", ha="center", fontsize=9)
ax.set_title("Average Default Rate by RFM Segment", fontsize=13, fontweight="bold")
ax.set_ylabel("Default rate (%)")
ax.tick_params(axis='x', rotation=30)
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/03_default_rate_by_segment.png", dpi=150)
plt.close()

# 4. Avg CreditScore by segment
fig, ax = plt.subplots(figsize=(8,4.8))
p2 = profile.sort_values("AvgCreditScore", ascending=False)
bars = ax.bar(p2["Segment"], p2["AvgCreditScore"], color=[seg_colors[s] for s in p2["Segment"]])
for b, v in zip(bars, p2["AvgCreditScore"]):
    ax.text(b.get_x()+b.get_width()/2, v+8, f"{v:.0f}", ha="center", fontsize=9)
ax.set_title("Average Credit Score by RFM Segment", fontsize=13, fontweight="bold")
ax.set_ylabel("Avg Credit Score (300-900 scale)")
ax.tick_params(axis='x', rotation=30)
ax.spines[["top","right"]].set_visible(False)
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/04_creditscore_by_segment.png", dpi=150)
plt.close()

# 5. RFM heatmap (R vs F, count of customers, sized bubble or heat)
fig, ax = plt.subplots(figsize=(6.5,5.5))
pivot = rfm.pivot_table(index="R_Score", columns="F_Score", values="CustomerID", aggfunc="count", fill_value=0)
pivot = pivot.sort_index(ascending=False)
im = ax.imshow(pivot.values, cmap="YlOrRd", aspect="auto")
ax.set_xticks(range(len(pivot.columns))); ax.set_xticklabels(pivot.columns)
ax.set_yticks(range(len(pivot.index))); ax.set_yticklabels(pivot.index)
for i in range(pivot.shape[0]):
    for j in range(pivot.shape[1]):
        ax.text(j, i, pivot.values[i,j], ha="center", va="center", fontsize=9,
                 color="white" if pivot.values[i,j] > pivot.values.max()/2 else "black")
ax.set_xlabel("Frequency Score")
ax.set_ylabel("Recency Score")
ax.set_title("Customer Density: Recency Score x Frequency Score", fontsize=12, fontweight="bold")
plt.colorbar(im, ax=ax, label="Customers")
plt.tight_layout()
plt.savefig(f"{CHART_DIR}/05_rf_heatmap.png", dpi=150)
plt.close()

print("Charts created:")
import os
for f in sorted(os.listdir(CHART_DIR)):
    print(" -", f)
