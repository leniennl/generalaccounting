import pandas as pd
import numpy as np
import os, sys
import xlwings as xw
import datetime as dt
from tkinter import Tk  # py3k


def change_revenue_group(group_number):
    if group_number == 10:
        return "111"
    elif group_number == 20:
        return "121"


# get path from clipboard
Filepath = Tk().selection_get(selection="CLIPBOARD")
if os.path.isdir(Filepath) == False:
    print("copy a valid path to clipboard")
    sys.exit()
counter = 0
for filefolders, subfolders, filenames in os.walk(Filepath):
    for filename in filenames:
        if str(filename).lower().startswith("status 571") or str(
            filename
        ).lower().startswith("571"):
            status_571_filepath = str(filefolders + "/" + filename)
            counter += 1
if counter == 0:
    print("No qualified file in this folder.")
    sys.exit()
elif counter >= 2:
    print(str(counter) + " qualified files found! Delete all but one.")
    sys.exit()

df_status = pd.read_excel(
    status_571_filepath,
    "Sheet2",
    index_col=None,
)

# check for and delete rows with "" in 2nd Item Number
index_row_to_delete = df_status[
    (df_status["2nd Item Number"] == "CONSULTING_STANDARD")
].index
df_status = df_status.drop(index_row_to_delete)  # drop  rows

# do a piovt
pivot = pd.pivot_table(
    df_status,
    values="Extended Amount",
    index=["Branch/Plant", "Revenue Group"],
    aggfunc=np.sum,
)
pivot = pivot.reset_index()
pivot["Revenue Group"].astype("str")
pivot["Branch/Plant"].astype("str")

# add pivot to existing one
with pd.ExcelWriter(
    status_571_filepath, engine="openpyxl", mode="a", if_sheet_exists="new"
) as writer:

    pivot.to_excel(
        excel_writer=writer,
        sheet_name="pivot",
        index=False,
    )

# add journals
# open raw data file
wb = xw.Book(status_571_filepath)
sht = wb.sheets["pivot"]
sht.activate()

# find the numbers of columns and rows in the sheet
num_col = sht.range("A1").end("right").column
num_row = sht.range("A1").end("down").row

sht.range((num_row + 2, num_col + 2)).number_format = "@"
sht.range((num_row + 2, num_col + 2)).value = "112.1060"
sht.range((num_row + 2, num_col + 3)).formula = "=SUM(c2:c" + str(num_row) + ")"
sht.range((num_row + 2, num_col + 3)).value = sht.range(
    (num_row + 2, num_col + 3)
).value
sht.range((num_row + 2, num_col + 9)).value = "Status 571 Rev Accrual"

for i in range(2, num_row + 1):
    sht.range((num_row + 1 + i, num_col + 2)).number_format = "@"
    sht.range((num_row + 1 + i, num_col + 2)).value = (
        str(sht.range((i, 1)).value)
        + str(change_revenue_group(sht.range((i, 2)).value))
    ).replace(".", "") + ".3100"
    sht.range((num_row + 1 + i, num_col + 4)).value = sht.range((i, 3)).value
    sht.range((num_row + 1 + i, num_col + 9)).value = "Status 571 Rev Accrual"

wb.save()

print("\n\nTo Check Total, & Copy Range to Clipboard!\n\n")
