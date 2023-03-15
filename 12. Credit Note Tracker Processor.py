#! Python3
# to find net-off-able lines in 112.2055 GRNV file: whole PO, or partial PO
# copy path to the GRNV file in clipboard and run this script

import re
import xlwings as xw
from tkinter import Tk  # py3k
import os, sys, time
from PyPDF2 import PdfReader
import getCustomerState

# get path from clipboard, & find GRNV file
CreditNoteTrackerFilepath = Tk().selection_get(selection="CLIPBOARD")

if os.path.isdir(CreditNoteTrackerFilepath) == False:
    print("copy a valid path to Credit Note Tracker file to the clipboard")
    sys.exit()

counter = 0
for filefolders, subfolders, filenames in os.walk(CreditNoteTrackerFilepath):
    for filename in filenames:
        if str(filename).startswith("Credit Note Tracker"):
            CreditNoteTrackerFile = str(filefolders + "/" + filename)
            counter += 1

if counter == 0:
    print("No Credit Note Tracker file in this folder.")
    sys.exit()
elif counter >= 2:
    print(str(counter) + " Credit Note Tracker files found! Delete all but one.")
    sys.exit()


# prep
workingsheet = xw.Book(CreditNoteTrackerFile).sheets["Use this tab please"].copy()
workingsheet.name = "Result"

# remove filter, if any
if workingsheet.api.AutoFilterMode == True:
    workingsheet.api.AutoFilter.ShowAllData()

# RE for credit note numbers entered
creditnotenumber = re.compile(r"(CN|Credit)(#)*(\s)*\d{6}")

# delete each unnecessary rows
starttime = time.time()
for i in range(workingsheet.used_range.last_cell.row + 1, 1, -1):

    print(
        str(
            round(
                (workingsheet.used_range.last_cell.row - i)
                * 100
                / workingsheet.used_range.last_cell.row,
                1,
            )
        )
        + " % completed..."
    )

    if workingsheet.range(i, 7).value in ["Processed", "processed"]:
        workingsheet.range("A" + str(i), "M" + str(i)).delete(shift="up")
        continue
    elif (
        "invoice" in str(workingsheet.range(i, 6).value).lower()
    ):  # delete reinvoice rows
        workingsheet.range("A" + str(i), "M" + str(i)).delete(shift="up")
        continue
    elif (
        workingsheet.range(i, 10).value != None
        and creditnotenumber.search(str(workingsheet.range(i, 10).value)) != None
    ):  # delete rows with completed credit note number
        workingsheet.range("A" + str(i), "M" + str(i)).delete(shift="up")
    elif (
        workingsheet.range(i, 5).value is None
    ):  # delete rows with no credit amount
        workingsheet.range("A" + str(i), "M" + str(i)).delete(shift="up")

endtime1 = time.time()

print("Lines deleted... Time taken: " + str(endtime1 - starttime) + " seconds.")


# do journals
lastrow = workingsheet.range("A" + str(workingsheet.cells.last_cell.row)).end("up").row

runningtotal = 0
for j in range(2, lastrow + 1, 1):
    workingsheet.range("e" + str(lastrow + 3 + j)).number_format = "###0.0000"
    workingsheet.range("e" + str(lastrow + 3 + j)).value = str(getCustomerState.getCustomerState(int(str(workingsheet.range("c" + str(j)).value).strip(".0"))))+"0111.3100"
    workingsheet.range("f" + str(lastrow + 3 + j)).value = workingsheet.range(
        "e" + str(j)
    ).value
    workingsheet.range("f" + str(lastrow + 3 + j)).number_format = "#,##0.00"
    workingsheet.range("L" + str(lastrow + 3 + j)).value = workingsheet.range(
        "b" + str(j)
    ).value
    try:
        runningtotal += float(workingsheet.range("e" + str(j)).value)
    except:
        continue

workingsheet.range("e" + str(2 * lastrow + 3 + 1)).value = "'112.1070"
workingsheet.range("g" + str(2 * lastrow + 3 + 1)).value = runningtotal
workingsheet.range("L" + str(2 * lastrow + 3 + 1)).value = "Credit Note Tracker"

print("journal added. You may need to delete more rows")

endtime2 = time.time()

print("Time taken: " + str(endtime2 - endtime1) + " seconds.")
