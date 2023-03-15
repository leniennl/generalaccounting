# complete code

import pandas

df_catDB = pandas.read_excel(
    r"J:\Financial Reporting\FY23\04 Jan 2023\Purchase Inquiry - Corporate - Jan23 Result version 2.xlsx",
    "Category Codes DB (2)",
    index_col=None,
)
# df_catDB=df_catDB.drop(df_catDB.columns[1], axis=1)

df_poInq = pandas.read_excel(
    r"C:\Users\matthew.lee\Desktop\Purchase Inquiry - Corporate - Jul22 Report.xlsx",
    "Po Inq Jul22",
    index_col=None,
)
df_poInq = df_poInq.drop(df_poInq.columns[[9, 10, 11, 12, 13, 14, 15, 19, 20]], axis=1)

# replace received Qty 0 with 1
df_poInq["Quantity\nReceived"].replace(0.00, 1, inplace=True)

# add 3 Columns
# df_poInq .insert(12, "Category",None)
# df_poInq .insert(13, "Sub-Category",None)
df_poInq.insert(12, "Currency", "AUD")
df_poInq["Unit Price"] = df_poInq.apply(
    lambda row: row["Amount\nReceived"] / row["Quantity\nReceived"], axis=1
)

# get suupplier_number & Name
df_poInq = df_poInq.join(
    df_poInq["Supplier\nNumber"]
    .str.split(" - ", expand=True)
    .add_prefix("Supplier_Number_Name_")
)

# re-sort column orders
neworder = [
    "Company",
    "Order\nNumber",
    "Received Date",
    "Supplier_Number_Name_0",
    "2nd\nItem\nNumber",
    "Description",
    "Quantity\nReceived",
    "Unit Price",
    "Amount\nReceived",
    "Currency",
    "Unit of Measure",
    "Supplier_Number_Name_1",
    "Description\n2",
    "3rd\nItem\nNumber",
]
df_poInq = df_poInq[neworder]


# rename columns
df_poInq.rename(
    columns={
        "Order\nNumber": "PO#",
        "Supplier_Number_Name_0": "Local Supplier Number",
        "2nd\nItem\nNumber": "Local Item Number",
        "Description": "Item Description",
        "Supplier_Number_Name_1": "Supplier Name",
        "Description\n2": "Supplier Item Description",
        "3rd\nItem\nNumber": "Supplier Item Code",
    },
    inplace=True,
)

# write to excel
left_join = pandas.merge(df_poInq, df_catDB, on="Local Item Number", how="left")

left_join.to_excel(r"C:\Users\matthew.lee\Desktop\Jul2022.xlsx", index=False)
