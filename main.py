import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")


for filepath in filepaths:
    df = pd.read_excel(filepath)
    pdf = FPDF(orientation='P', unit="mm", format="A4")
    pdf.add_page()

# Printing out the page header
    page_name = filepath.removeprefix("invoices\\").removesuffix(".xlsx")
    pdf.set_font(family="Times", size=17, style="B")
    pdf.cell(w=0, h=12, txt=f"Invoice nr.{page_name[0:5]}", border=1)
    pdf.ln(12)
    pdf.cell(w=0, h=12, txt=f"Date {page_name[6:]}", border=1)
    print(page_name)

    pdf.output("output.pdf")
