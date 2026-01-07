# ===============================
# superstore_analysis.py / .ipynb
# ===============================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# -----------------------------
# Step 1: Load the dataset
# -----------------------------
DATA_PATH = r"C:\\Users\\enves\\Desktop\\Python\\Superstore.csv"  # change path according to your PC
df = pd.read_csv(DATA_PATH, encoding='latin1')

# View the first 5 rows
print(df.head())

# -----------------------------
# Step 2: Inspect the dataset
# -----------------------------
print(df.info())
print(df.isnull().sum())
print(df.describe())

# -----------------------------
# Step 3: Total, maximum, minimum profit
# -----------------------------
total_profit = df['Profit'].sum()
max_profit = df['Profit'].max()
min_profit = df['Profit'].min()
mean_profit = df['Profit'].mean()
std_profit = df['Profit'].std()

print(f"Total Profit: {total_profit}")
print(f"Max Profit: {max_profit}")
print(f"Min Profit: {min_profit}")
print(f"Mean Profit: {mean_profit}")
print(f"Std Profit: {std_profit}")

# -----------------------------
# Step 4: Losses by Category and Sub-Category
# -----------------------------
category_loss = df.groupby('Category')['Profit'].sum().sort_values()
print(f"Category with largest loss: {category_loss.index[0]}, Loss: {category_loss.iloc[0]}")

subcat_loss = df.groupby('Sub-Category')['Profit'].sum().sort_values()
print(f"Sub-Category with largest loss: {subcat_loss.index[0]}, Loss: {subcat_loss.iloc[0]}")

# -----------------------------
# Step 5: Visualizations
# -----------------------------
# Profit by Category
category_profit = df.groupby('Category')['Profit'].sum()
category_profit.plot(kind='bar', figsize=(8,5), title='Total Profit per Category', color='skyblue')
plt.ylabel('Profit')
plt.tight_layout()
plt.show()

# Profit by Sub-Category
subcat_profit = df.groupby('Sub-Category')['Profit'].sum().sort_values()
subcat_profit.plot(kind='barh', figsize=(8,6), title='Total Profit per Sub-Category', color='lightgreen')
plt.xlabel('Profit')
plt.tight_layout()
plt.show()

# -----------------------------
# Step 6: Export Excel report (summary)
# -----------------------------
summary = pd.DataFrame({
    'Total Sales': df.groupby('Category')['Sales'].sum(),
    'Total Profit': df.groupby('Category')['Profit'].sum()
}).reset_index()

subcat_summary = df.groupby(['Category','Sub-Category'])[['Sales','Profit']].sum().reset_index()

# Output folder (relative path)
output_folder = os.path.join(os.getcwd(), "outputs")
charts_folder = os.path.join(output_folder, "charts")

os.makedirs(charts_folder, exist_ok=True)

# Save Excel files (make sure openpyxl is installed)
summary.to_excel(os.path.join(output_folder, "Category_Summary.xlsx"), index=False)
subcat_summary.to_excel(os.path.join(output_folder, "SubCategory_Summary.xlsx"), index=False)

print("Reports saved successfully in 'outputs' folder!")