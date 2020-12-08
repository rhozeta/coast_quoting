import pandas as pd

def getAllProducts():
    df_full = pd.read_csv("product_data_clean.csv")
    return (df_full)

if "NS744" in str(getAllProducts()["SKU"]):
    print("yes")
else:
    print("no")