import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
import os

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    # Installed openpyxl module to read data from Excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Extracting filename and invoice from paths
    filename = Path(filepath).stem
    #Extracting invoice from filename and splitting at date
    invoice_no = filename.split("-")[0]
    date = filename.split("-")[1]

    #Create PDF file from Excel file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_no}", ln=1)

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    #Create PDF that creates dynamic name automatically
    pdf.output(f"PDFs/{filename}.pdf")