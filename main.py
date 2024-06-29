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
    invoices_nr = filename.split("-")[0]
    pdf.set_font(family="Times",style="B",size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice no: {invoices_nr}", align="L")
    pdf.output(f"PDFs/{filename}.pdf")


