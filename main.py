import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    # Installed openpyxl module to read data from Excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Extracting filename and invoice from paths
    filename = Path(filepath).stem
    #Extracting invoice and date from filename
    invoice_no, date = filename.split("-")

    # Create PDF file from Excel file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_no}", ln=1)

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Table Header by using pandas
    header = df.columns
    # List Comprehension to remove underscores in header
    header = [item.replace("_", " ").title() for item in header]
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(w=30, h=8, txt=header[0], border=1)
    pdf.cell(w=70, h=8, txt=header[1], border=1)
    pdf.cell(w=30, h=8, txt=header[2], border=1)
    pdf.cell(w=30, h=8, txt=header[3], border=1)
    pdf.cell(w=30, h=8, txt=header[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="B", size=10)
        pdf.set_text_color(40, 40, 40)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Create PDF that creates dynamic name automatically
    pdf.output(f"PDFs/{filename}.pdf")