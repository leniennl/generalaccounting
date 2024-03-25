#!/usr/bin/env python3
# extract ending balance from bank statement, pass to clipboard

import pdfplumber
import re
import pyperclip
import sys


def extract_Text_From_Pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text()
        return text
    except Exception as e:
        print("An error occurred with PDF text extraction:", e)
        return None


def get_Bank_Closing_Balance(text):
    pattern = r"Closing balance:\s+AUD\s+([\d,]+\.\d+)\+"
    match = re.search(pattern, text)
    if match:
        extracted_float = float(match.group(1).replace(",", ""))
        return (extracted_float)
    else:
        print("Pattern not found.")
        return None


def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <filePath>")
        sys.exit()
    filePath = sys.argv[1]

    extracted_text = extract_Text_From_Pdf(filePath)
    bank_closing_balance = get_Bank_Closing_Balance(extracted_text)
    if extracted_text is not None:
        if bank_closing_balance is not None:
            pyperclip.copy(bank_closing_balance)
        else:
            print("get_Bank_Closing_Balance method failed.")
            pyperclip.copy("")
    else:
        print("extract_Text_From_Pdf method failed.")
        pyperclip.copy("")

if __name__ == "__main__":
    main()
