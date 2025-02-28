import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


#pandas need openpyxl to read excel files
#xlsx is a used for reading excel sheets

filepaths = glob.glob("invoices/*.xlsx")
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoices_nr = filename.split("-")[0]
    pdf.set_font(family="Times",size=16,style="B")
    pdf.cell(w=50 , h = 8 , txt=f"Invoice nr,{invoices_nr}")
    pdf.output(f"PDFS/{filename}.pdf")
