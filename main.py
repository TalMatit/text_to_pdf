import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")


for filepath in filepaths:
    df = pd.read_excel(filepath)
    pdf = FPDF(orientation='P', unit="mm", format="A4")
    pdf.add_page()

# Printing out the page header
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Invoice nr.{invoice_nr}", border=0)
    pdf.ln(8)
    pdf.cell(w=0, h=8, txt=f"Date {invoice_date}", border=0)

    pdf.output(f"PDFs/{filename}.pdf")
