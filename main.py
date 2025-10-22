import pandas as pd
import matplotlib.pyplot as plt
import os

# Paths
input_path = "data/Customer_Report.xlsx"
output_dir = "outputs"
os.makedirs(output_dir, exist_ok=True)

# Load data
df = pd.read_excel(input_path)
print("Data Loaded. Shape:", df.shape)

# Define months
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
spend_cols = [f"Spend_{m}" for m in months]
repay_cols = [f"Repayment_{m}" for m in months]

# Clean data
df.fillna(0, inplace=True)
df = df[df["Age"] >= 18]

# Calculate totals
df["Total_Spend"] = df[spend_cols].sum(axis=1)
df["Total_Repayment"] = df[repay_cols].sum(axis=1)
df["Outstanding"] = df["Total_Spend"] - df["Total_Repayment"]
df["Profit"] = df["Outstanding"] * 0.029  # 2.9% interest on due
df["OverLimit_Months"] = sum([df[f"Spend_{m}"] > df["CreditLimit"] for m in months])

# Group-level summaries
segment_summary = df.groupby("CardType")[["Total_Spend", "Total_Repayment", "Outstanding", "Profit"]].mean()
age_summary = df.groupby("AgeGroup")[["Total_Spend", "Total_Repayment", "Profit"]].mean()

# Export Excel outputs
df.to_excel(f"{output_dir}/cleaned_customer_data.xlsx", index=False)
segment_summary.to_excel(f"{output_dir}/segment_summary.xlsx")
age_summary.to_excel(f"{output_dir}/age_summary.xlsx")

# Visuals
plt.figure(figsize=(8,5))
segment_summary["Total_Spend"].plot(kind="bar", color="skyblue", title="Average Spend by Card Type")
plt.ylabel("Average Spend")
plt.tight_layout()
plt.savefig(f"{output_dir}/segment_spend_chart.png")
plt.show()

print("✅ Analysis complete. Files saved in 'outputs' folder.")
