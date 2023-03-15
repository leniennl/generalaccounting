# complete code
import xlwings as xw
import pandas
import numpy as np
import datetime as dt

POdf = pandas.read_excel(
    r"J:\Financial Reporting\FY23\04 Jan 2023\Purchase Inquiry - Corporate - Jan23 Data.xlsx",
    "PO Inq Jan23",
    index_col=None,
)
GLdf = pandas.read_excel(
    r"J:\Financial Reporting\FY23\04 Jan 2023\Purchase Inquiry - Corporate - Jan23 Data.xlsx",
    "GL Inq Jan23",
    index_col=None,
)

address_book = pandas.read_excel(
    r"J:\Financial Reporting\FY23\04 Jan 2023\Purchase Inquiry - Corporate - Jan23 Result.xlsx",
    "Add Book",
    index_col=None,
)
supplier_category = pandas.read_excel(
    r"J:\Financial Reporting\FY23\04 Jan 2023\Purchase Inquiry - Corporate - Jan23 Result.xlsx",
    "Supplier Category",
    index_col=None,
)
supplier_cat_code = pandas.read_excel(
    r"J:\Financial Reporting\FY23\04 Jan 2023\Purchase Inquiry - Corporate - Jan23 Result.xlsx",
    "Supplier Cat Code",
    index_col=None,
)
address_book = address_book.drop(address_book.columns[[1, 2, 3, 4, 5, 7]], axis=1)
supplier_category = supplier_category.drop(supplier_category.columns[[2, 5]], axis=1)

# work on GL Inquiry tab
GL_no_po_df = GLdf.loc[GLdf["Purchase Order"].notnull() == False]  # no PO lines

indexAge = GL_no_po_df[
    (GL_no_po_df["Document\nType"] == "PT")
    | (GL_no_po_df["Document\nType"] == "PN")
    | (GL_no_po_df["Address\nNumber"] == "Grand Total")
].index
GL_no_po_df.drop(indexAge, inplace=True)  # drop PT & PN & Total rows

GL_no_po_df = GL_no_po_df.filter(
    ["Address\nNumber", "JE\nExplanation", "GL\nAmount", "Exchange\nRate", "GL Date"],
    axis=1,
)  # rename column name

GL_no_po_df.columns=['Supplier_Number_Name_0', 'Supplier_Number_Name_1', 'Amount\nReceived',  'Exchange\nRate', 'Received Date']

df_splitSupplierNumber = POdf.join(
    POdf["Supplier\nNumber"]
    .str.split(" - ", expand=True)
    .add_prefix("Supplier_Number_Name_")
)  # split supplier # & name

# add GL Inq & PO Inq
df_combined = pandas.concat([df_splitSupplierNumber, GL_no_po_df], axis=0)
df_combined["Supplier_Number_Name_0"] = df_combined["Supplier_Number_Name_0"].astype(
    int
)

# look up vendor type
df_lookupvendortype = pandas.merge(
    df_combined, address_book, on="Supplier_Number_Name_0", how="left"
)

# exit codes if vendor type is NA

# delete vendor type except V
rowstodrop = df_lookupvendortype["Sch Typ"].isin(
    ["O", "E", "CX", "CI", "CP", "VI", "X", "TAX", "GD"]
)
df_lookupvendortype = df_lookupvendortype.drop(df_lookupvendortype[rowstodrop].index)

# generate PO Inquiry Summary tab

# do a piviot table on df_lookupvendortype to tally Amount
pivot = pandas.pivot_table(
    df_lookupvendortype,
    values="Amount\nReceived",
    index="Supplier_Number_Name_0",
    aggfunc=np.sum,
)
pivot["Local supplier code"] = pivot.index

# month & year
# month_report = str(df_lookupvendortype["Received Date"][22].date().month)
month_report = df_lookupvendortype["Received Date"][22].date().strftime('%b')
year_report = str(df_lookupvendortype["Received Date"][22].date().year)

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

# rename
pivot_merge.rename(
    columns={
        "Amount\nReceived": "Amount",
        "Local supplier code": "Supplier Number",
        "Local supplier name": "Supplier Name",
        "Supplier Category": "Category of Spend",
    },
    inplace=True,
)

# sort column
neworder = [
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
pivot_merge = pivot_merge[neworder]

pivot_merge.to_excel(r"C:\Users\matthew.lee\Desktop\Purchase Inquiry.xlsx", index=False)

# tidy up workbook


with xw.App(visible=False) as app:
    wb = xw.Book(r"C:\Users\matthew.lee\Desktop\Purchase Inquiry.xlsx")

    for ws in wb.sheets:
        ws.autofit(axis="columns")

    wb.save("C:\Users\matthew.lee\Desktop\Purchase Inquiry.xlsx")
    wb.close()
