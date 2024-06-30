import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename0 = Path(filepath).stem
    filename = filename0.split(" ")[1]
    invoices_nr, date = filename.split("-")
    pdf.set_font(family="Times",style="B",size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice no: {invoices_nr}", align="L",ln =1)
    pdf.set_font(family="Times",style="B",size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", align="L",ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #For header
    column = df.columns
    column = [item.replace("_"," ").title() for item in column]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=column[0], border=1)
    pdf.cell(w=40, h=8, txt=column[1], border=1)
    pdf.cell(w=40, h=8, txt=column[2], border=1)
    pdf.cell(w=30, h=8, txt=column[3], border=1)
    pdf.cell(w=30, h=8, txt=column[4], border=1,ln=1)


    #for row

    for index, row in df.iterrows():

        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]),border=1)
        pdf.cell(w=35, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]),border=1,ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30,h=12,txt=f"The Total Price is {total_sum}",ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=25, h=12, txt=f"Aryan Kanada")
    pdf.image("Photo.png",w=20, h=20)

    pdf.output(f"PDFs/{filename}.pdf")

