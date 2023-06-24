import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    # The above line can be written as below
    # invoice_nr = filename.split("-")[0]
    # date = filename.split("-")[1]

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=10, txt=f"Invoice nr: {invoice_nr}", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=10, txt=f"Date : {date}",ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = df.columns

    # Rename column names into readable format
    columns_new = [column.replace("_", " ").title() for column in columns]

    # Add header columns
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=columns_new[0], border=1)
    pdf.cell(w=60, h=8, txt=columns_new[1], border=1)
    pdf.cell(w=40, h=8, txt=columns_new[2], border=1)
    pdf.cell(w=30, h=8, txt=columns_new[3], border=1)
    pdf.cell(w=30, h=8, txt=columns_new[4], border=1, ln=1)

    for index, row in df.iterrows():
        # Add table columns
        pdf.set_font(family="Times", style="B", size=10)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # Row with total price
    total_price = df["total_price"].sum()
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=0, h=8, txt=f"The total price is {total_price}", ln=1)
    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=35, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
