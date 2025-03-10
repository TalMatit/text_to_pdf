import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
import pandas as pd

filepaths = glob.glob("invoices/*.xlsx")



for filepath in filepaths:
    # adding a new page
    pdf = FPDF(orientation='P', unit="mm", format="A4")
    pdf.add_page()

    # assigning variables to the information in the filen_name
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")

    # Printing out the page header
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Invoice nr.{invoice_nr}", border=0, ln=1)

    # Printing out the date
    pdf.cell(w=0, h=8, txt=f"Date {invoice_date}", border=0, ln=1)

    ln = 5
    # Reading the excel file and producing an output on the page
    df = pd.read_excel(filepath)
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    # Adding the column names
    pdf.set_font(family="times", size=10, style="B")
    pdf.set_text_color(100,100,100)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], ln=1, border=1)

    for index, row in df.iterrows():
        pdf.set_font(family="times", size=12, style="")
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=12, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=12, txt=row["product_name"], border=1)
        pdf.cell(w=30, h=12, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=12, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=12, txt=str(row["total_price"]), ln=1, border=1)

    # Adding the total sum cell
    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=12, txt="", border=1)
    pdf.cell(w=70, h=12, txt="", border=1)
    pdf.cell(w=30, h=12, txt="", border=1)
    pdf.cell(w=30, h=12, txt="", border=1)
    pdf.cell(w=30, h=12, txt=str(total_sum), border=1)

    # Add total sum sentence
    pdf.cell(w=30, h=12, txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.cell(w=25, h=12, txt=f"Pythonhow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
