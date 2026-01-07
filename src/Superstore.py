# ===============================
# analysis.py - Final Version
# ===============================

import pandas as pd
import os

# ----  Set dataset path ----
DATA_PATH = r"C:\\Users\\enves\\Desktop\\Python\\superstore.csv"  # change according to your location

# ----  Load the dataset ----
df = pd.read_csv(DATA_PATH, encoding='latin1')

# ----  Functions for summary ----
def summary_category(df):
    """
    Creates summary for Category
    """
    return df.groupby('Category')[['Sales','Profit']].sum().reset_index()

def summary_subcategory(df):
    """
    Creates summary for Sub-Category
    """
    return df.groupby(['Category','Sub-Category'])[['Sales','Profit']].sum().reset_index()

# ---  Create output folder ----
output_folder = os.path.join(os.getcwd(), "outputs")
os.makedirs(output_folder, exist_ok=True)

# ---- Save reports to Excel ----
summary_category(df).to_excel(os.path.join(output_folder,"Category_Summary.xlsx"), index=False)
summary_subcategory(df).to_excel(os.path.join(output_folder,"SubCategory_Summary.xlsx"), index=False)

print("Reports were successfully saved in the 'outputs/' folder!")
