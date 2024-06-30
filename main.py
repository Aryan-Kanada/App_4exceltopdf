import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename0 = Path(filepath).stem
    filename = filename0.split(" ")[1]
    invoices_nr, date = filename.split("-")
    pdf.set_font(family="Times",style="B",size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice no: {invoices_nr}", align="L",ln=1)
    pdf.set_font(family="Times",style="B",size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", align="L",ln=1)
    pdf.output(f"PDFs/{filename}.pdf")



