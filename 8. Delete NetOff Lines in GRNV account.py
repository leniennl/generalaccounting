#! Python3
# to find net-off-able lines in 112.2055 GRNV file: whole PO, or partial PO
# copy path to the GRNV file in clipboard and run this script

import pandas as pd
import xlwings as xw
from tkinter import Tk  # py3k
import os, sys


# get path from clipboard, & find GRNV file
GRNVpath = Tk().selection_get(selection="CLIPBOARD")

if os.path.isdir(GRNVpath) == False:
    print("copy a valid path to GRNV file to the clipboard")
    sys.exit()

counter = 0
for filefolders, subfolders, filenames in os.walk(GRNVpath):
    for filename in filenames:
        if str(filename).startswith("112.2055 GRNV"):
            GRNVfile = str(filefolders + "/" + filename)
            counter += 1

if counter == 0:
    print("No GRNV in this folder.")
    sys.exit()
elif counter >= 2:
    print(str(counter) + " GRNV files found! Delete all but one.")
    sys.exit()


# find last row in column A of Worksheet " 2055- GRNV "
wb = xw.Book(GRNVfile)
LastRow = (
    wb.sheets["2055 - GRNV"]
    .range("A" + str(wb.sheets["2055 - GRNV"].cells.last_cell.row))
    .end("up")
    .row
)
wb.close()


# read GRNV data into panda dataframe
df = pd.read_excel(GRNVfile, sheet_name="2055 - GRNV", nrows=LastRow, usecols="A:ar")

# find matching lines
for po in df["Purchase Order"].unique():
    if po == "NaN":
        continue

    elif (
        abs(df.loc[df["Purchase Order"] == po]["LT 1 Amount"].sum()) < 1e-10
    ):  # whole PO = 0
        df.loc[df["Purchase Order"] == po, "To Delete PO"] = "cccc"

    else:
        # find netting off LT 1 Amt lines' relative position in its PO
        makeaseries = pd.Series(df.loc[df["Purchase Order"] == po]["LT 1 Amount"])
        netoffposition = []
        for firstnumberposition in range(1, makeaseries.size):
            if firstnumberposition in netoffposition:
                continue  # skip to the next iteration when the number has already been marked and stored the list(netoffposition)
            for secondnumberposition in range(
                firstnumberposition + 1, makeaseries.size + 1
            ):
                if secondnumberposition in netoffposition:
                    continue  # skip to the next iteration when the number has already been marked and stored in the list(netoffposition)
                if (
                    makeaseries.iloc[firstnumberposition - 1]
                    + makeaseries.iloc[secondnumberposition - 1]
                    == 0
                ):
                    netoffposition.append(firstnumberposition)
                    netoffposition.append(secondnumberposition)
                    break  # leave inner loop, return to outer loop

        # change netoffposition to an index number (relative postion to absolute position)
        abosoluteindexposition = []
        for i in netoffposition:
            abosoluteindexposition.append(
                df.loc[df["Purchase Order"] == po].iloc[i - 1].name
            )

        # Update "To Delete PO " column value as per absolute index position
        for k in abosoluteindexposition:
            df.at[k, "To Delete PO"] = "dddd"


# export result to local file and clipboard
df.to_excel(GRNVpath + "/" + "grnvtoDelete " + GRNVfile[-13:-5] + ".xlsx", index=False)
print(
    f'Lean Version Exported in the GRNV folder.\n"To Delete PO" Column also pasted to clipboard'
)
df["To Delete PO"].to_clipboard(excel=True, sep=None, index=False)
