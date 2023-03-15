#! python 3
# process PO Inquiry Data file to produce 2 tables for Corporate

import xlwings as xw
import pandas
import numpy as np
import datetime as dt
import sys, os
from tkinter import Tk  # py3k

# get path from clipboard
Filepath = Tk().selection_get(selection="CLIPBOARD")
if os.path.isdir(Filepath) == False:
    print("copy a valid path to PO file to the clipboard")
    sys.exit()
counter = 0
for filefolders, subfolders, filenames in os.walk(Filepath):
    for filename in filenames:
        if str(filename).startswith("Purchase Inquiry"):
            POfilepath = str(filefolders + "/" + filename)
            counter += 1
if counter == 0:
    print("No PO file in this folder.")
    sys.exit()
elif counter >= 2:
    print(str(counter) + " PO files found! Delete all but one.")
    sys.exit()

# find month & year in names as variable
namepart = (
    POfilepath[POfilepath.find("Corporate") + 9 : POfilepath.find("Data")]
    .strip()
    .strip("-")
    .strip()
)
# read in tables
PO_DB_PATH = "J:\\Financial Reporting\\FY23\\05 Feb 2023\\PO Inquiry DataBase.xlsx"
POdf = pandas.read_excel(
    POfilepath,
    "PO Inq " + namepart,
    index_col=None,
)
GLdf = pandas.read_excel(
    POfilepath,
    "GL Inq " + namepart,
    index_col=None,
)
df_catDB = pandas.read_excel(
    PO_DB_PATH,
    "Category Codes DB",
    index_col=None,
)
address_book = pandas.read_excel(
    PO_DB_PATH,
    "Add Book",
    index_col=None,
)
supplier_category = pandas.read_excel(
    PO_DB_PATH,
    "Supplier Category",
    index_col=None,
)
supplier_cat_code = pandas.read_excel(
    PO_DB_PATH,
    "Supplier Cat Code",
    index_col=None,
)
df_catDB = df_catDB.drop(df_catDB.columns[1], axis=1)
address_book = address_book.drop(address_book.columns[[1, 2, 3, 4, 5, 7]], axis=1)
supplier_category = supplier_category.drop(supplier_category.columns[[2, 5]], axis=1)

# ...................Finance Angel.................................................
# work on GL Inquiry tab
GL_no_po_df = GLdf.loc[GLdf["Purchase Order"].notnull() == False]  # no PO lines
index_row_to_delete = GL_no_po_df[
    (GL_no_po_df["Document\nType"] == "PT")
    | (GL_no_po_df["Document\nType"] == "PN")
    | (GL_no_po_df["Address\nNumber"] == "Grand Total")
].index
GL_no_po_df = GL_no_po_df.drop(index_row_to_delete)  # drop PT & PN & Total rows
GL_no_po_df = GL_no_po_df.filter(
    ["Address\nNumber", "JE\nExplanation", "GL\nAmount", "Exchange\nRate", "GL Date"],
    axis=1,
)  # rename column name to match GL Inq Table with PO Inq table
GL_no_po_df.columns = [
    "Supplier_Number_Name_0",
    "Supplier_Number_Name_1",
    "Amount\nReceived",
    "Exchange\nRate",
    "Received Date",
]
df_splitSupplierNumber = POdf.join(
    POdf["Supplier\nNumber"]
    .str.split(" - ", expand=True)
    .add_prefix("Supplier_Number_Name_")
)  # split supplier # & name
# vertically add GL Inq & PO Inq
df_combined = pandas.concat([df_splitSupplierNumber, GL_no_po_df], axis=0)
df_combined["Supplier_Number_Name_0"] = df_combined["Supplier_Number_Name_0"].astype(
    int
)
# look up vendor type, and delete vendor types, leaving type V ( and None)
df_lookupvendortype = pandas.merge(
    df_combined, address_book, on="Supplier_Number_Name_0", how="left"
)
rowstodrop = df_lookupvendortype["Sch Typ"].isin(
    ["O", "E", "CX", "CI", "CP", "VI", "X", "TAX", "GD"]
)
df_lookupvendortype = df_lookupvendortype.drop(df_lookupvendortype[rowstodrop].index)
# identify rows with missing vendor Category
df_lookupvendortype["Sch Typ"]=df_lookupvendortype["Sch Typ"].fillna('').replace('',np.nan)
df_missing_vendor_inf_in_db=df_lookupvendortype[df_lookupvendortype["Sch Typ"].isna()]
df_missing_vendor_inf_in_db=df_missing_vendor_inf_in_db[['Supplier_Number_Name_0',"Supplier_Number_Name_1"]]
# month & year
month_report = df_lookupvendortype["Received Date"][22].date().strftime("%b")
year_report = str(df_lookupvendortype["Received Date"][22].date().year)

# generate PO Inquiry Summary tab
# do a piviot table on df_lookupvendortype to tally Amount
pivot = pandas.pivot_table(
    df_lookupvendortype,
    values="Amount\nReceived",
    index="Supplier_Number_Name_0",
    aggfunc=np.sum,
)
pivot["Local supplier code"] = pivot.index  # make index(supplier code) a column
# add other necessary columns
pivot.insert(0, "Company", "112")
pivot.insert(2, "Month", month_report)
pivot.insert(3, "Year", year_report)
pivot.insert(4, "Currency", "AUD")
pivot.insert(5, "Approved by Purchasing Dept", "Approved")
# merge with Supplier Category table
pivot_merge = pandas.merge(
    pivot, supplier_category, on="Local supplier code", how="left"
)
# rename column names
pivot_merge.rename(
    columns={
        "Amount\nReceived": "Amount",
        "Local supplier code": "Supplier Number",
        "Local supplier name": "Supplier Name",
        "Supplier Category": "Category of Spend",
    },
    inplace=True,
)
# re-arrange column
NEWORDER = [
    "Company",
    "Supplier Number",
    "Month",
    "Year",
    "Amount",
    "Currency",
    "Supplier Name",
    "Category of Spend",
    "Direct_Indirect",
    "Approved by Purchasing Dept",
]
df_po_summary = pivot_merge[NEWORDER]  # re-arrange columns

# ...................Purchasing/Lab Angel.................................................
df_poInq = POdf
df_poInq = df_poInq.drop(df_poInq.columns[[9, 10, 11, 12, 13, 14, 15, 19, 20]], axis=1)
df_poInq["Quantity\nReceived"].replace(
    0.00, 1, inplace=True
)  # replace received Qty 0 with 1
df_poInq.insert(12, "Currency", "AUD")  # add 1 Column
df_poInq["Unit Price"] = df_poInq.apply(
    lambda row: row["Amount\nReceived"] / row["Quantity\nReceived"], axis=1
)  # add 1 Column
df_poInq = df_poInq.join(
    df_poInq["Supplier\nNumber"]
    .str.split(" - ", expand=True)
    .add_prefix("Supplier_Number_Name_")
)  # split suupplier_number & Name
NEWORDER_2 = [
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
]  # re-arrange column orders
df_poInq = df_poInq[NEWORDER_2]
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
)  # rename columns
df_po_lab = pandas.merge(df_poInq, df_catDB, on="Local Item Number", how="left")

# write to excel
with pandas.ExcelWriter(
    Filepath
    + "/112 PO ACC File "
    + month_report
    + " "
    + year_report
    + " Report.xlsx"
) as writer:
    df_po_summary.to_excel(
        excel_writer=writer,
        sheet_name="Supplier Summary " + month_report + " " + year_report,
        index=False,
    )
    df_po_lab.to_excel(
        excel_writer=writer,
        sheet_name="PO Inquiry " + month_report + " " + year_report,
        index=False,
    )
with pandas.ExcelWriter(
    Filepath
    + "/Purchase Inquiry File "
    + month_report
    + " "
    + year_report
    + " Workings.xlsx"
) as writer:
    df_lookupvendortype.to_excel(
        excel_writer=writer,
        sheet_name="Workings " + month_report + " " + year_report,
        index=False,
    )
    df_missing_vendor_inf_in_db=df_missing_vendor_inf_in_db.to_excel(
        excel_writer=writer,
        sheet_name="Missing These Vendor Info ",
        index=False,
    )

# autofit workbook
with xw.App(visible=False) as app:
    wb = xw.Book(
        Filepath
        + "/112 PO ACC File "
        + month_report
        + " "
        + year_report
        + " Report.xlsx"
    )
    for ws in wb.sheets:
        ws.autofit(axis="columns")
    wb.save()
    wb.close()

print(
    "Purchase Inquiry File Exported.\nCheck for Empty cells in Summary tab, & Update PO Database."
)
