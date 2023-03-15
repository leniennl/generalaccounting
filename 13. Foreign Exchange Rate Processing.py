import xlwings as xw
import sys, os
from tkinter import Tk  # py3k
import time
import webbrowser


# get path from clipboard
fxfilepath = Tk().selection_get(selection="CLIPBOARD")

if os.path.isdir(fxfilepath) == False:
    print("copy a valid path to clipboard")
    sys.exit()

filecounter = 0

for filefolders, subfolders, filenames in os.walk(fxfilepath):
    for filename in filenames:
        if "Fx Rate" in str(filename):
            fx_rate_wb = xw.Book(filefolders+"/"+ filename)
            filecounter += 1


if filecounter == 0:
    print("No FX filein this folder.")
    exit
else:
    print("FX file found.")

xw.App.visible=True
data_average_ws = fx_rate_wb.sheets[0]
data_average_ws.activate()
xw.Range("v:ar").api.Delete()
time.sleep(0.1)
xw.Range("r:t").api.Delete()
time.sleep(0.1)
xw.Range("j:p").api.Delete()
time.sleep(0.1)
xw.Range("d:h").api.Delete()
time.sleep(0.1)


last_row = (
    data_average_ws.range("B" + str(data_average_ws.cells.last_cell.row)).end("up").row
)
#  find last row of month in table
for i in range(8, last_row - 1):
    if data_average_ws.range((i, 3)).value == None:
        month_of_fx = i - 1
        break

# Workout fx rate
fx_aud = 1 / data_average_ws.range((month_of_fx, 3)).value
fx_cad = fx_aud * data_average_ws.range((month_of_fx, 4)).value
fx_cny = fx_aud * data_average_ws.range((month_of_fx, 5)).value
fx_eur = fx_aud * data_average_ws.range((month_of_fx, 6)).value


#  write to cells
data_average_ws.range((last_row + 8, 3)).value = fx_aud
data_average_ws.range((last_row + 8, 4)).value = fx_cad
data_average_ws.range((last_row + 8, 5)).value = fx_cny
data_average_ws.range((last_row + 8, 6)).value = fx_eur
data_average_ws.range((last_row + 7, 7)).value = "AUD/NZD"
data_average_ws.range((last_row + 7, 8)).value = "https://www.google.com/finance/quote/AUD-NZD"



fx_rate_wb.save()

# data closing sheet
data_closing_ws = fx_rate_wb.sheets[1]
data_closing_ws.activate()
xw.Range("l:v").api.Delete()
time.sleep(0.1)
xw.Range("j:j").api.Delete()
time.sleep(0.1)
xw.Range("f:h").api.Delete()
time.sleep(0.1)
xw.Range("c:d").api.Delete()
time.sleep(0.1)

last_row = (
    data_closing_ws.range("a" + str(data_average_ws.cells.last_cell.row)).end("up").row
)

#  find last row of month in table
for i in range(7, last_row + 1):
    if data_closing_ws.range((i, 2)).value == None:
        month_of_fx = i - 1
        break

# Workout fx rate
fx_aud_closing = 1 / data_closing_ws.range((month_of_fx, 2)).value
fx_cad_closing = fx_aud * data_closing_ws.range((month_of_fx, 3)).value
fx_cny_closing = fx_aud * data_closing_ws.range((month_of_fx, 4)).value
fx_eur_closing = fx_aud * data_closing_ws.range((month_of_fx, 5)).value


#  write to cells
data_closing_ws.range((last_row + 3, 2)).value = fx_aud_closing
data_closing_ws.range((last_row + 3, 3)).value = fx_cad_closing
data_closing_ws.range((last_row + 3, 4)).value = fx_cny_closing
data_closing_ws.range((last_row + 3, 5)).value = fx_eur_closing
data_closing_ws.range((last_row + 2, 6)).value = "AUD/NZD"
data_average_ws.range((last_row + 2, 7)).value = "https://www.google.com/finance/quote/AUD-NZD"


fx_rate_wb.save()

print("FX file processed!")


webbrowser.open("https://www.google.com/finance/quote/AUD-NZD")
