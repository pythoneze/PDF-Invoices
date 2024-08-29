import pandas as pd
import glob as gb
from fpdf import FPDF
from pathlib import Path

FILEPATHS = gb.glob("invoices/*xlsx*")


def create_pdf(filename, invoice_number):
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_number}")
    pdf.output(f"PDFs/{filename}.pdf")


def proces_invoices():
    for filepath in FILEPATHS:
        df = pd.read_excel(filepath, sheet_name="Sheet 1")
        filename = Path(filepath).stem
        invoice_number = filename.split("-")[0]
        create_pdf(filename, invoice_number)


if __name__ == "__main__":
    proces_invoices()