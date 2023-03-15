from PyPDF2 import PdfReader
import re

STATES = {
    "NSW": "1202",
    "VIC": "1203",
    "QLD": "1204",
    "WA": "1206",
    "TAS": "TAS",
    "NT": "NT",
    "SA": "SA"
}


def getCustomerState(customer_number):

    address_re = re.compile(
        str(customer_number) + r"(.)*\n(.)*\n(.)*\n(.)*\n(.)*\n(.)*\n"
    )
    state_re = re.compile(r"(NSW|VIC|WA|TAS|NT|SA|QLD)")

    reader = PdfReader(
        r"C:\Users\matthew.lee\Dropbox\Side Hussle\Python\Work In Progress\who's who report.pdf"
    )
    # reader= PdfReader(r"C:\Users\lenie\Dropbox\Side Hussle\Python\Work In Progress\who's who report.pdf")

    NumPages = len(reader.pages)

    for i in range(0, NumPages):
        PageObj = reader.pages[i]
        Text = PageObj.extract_text()
        address_match = address_re.search(Text)
        if address_match is None:
            continue
        else:
            state_match = state_re.search(address_match.group())
            if state_match is None:
                continue
            else:
                return STATES[state_match.group()]



