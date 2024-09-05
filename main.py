import pandas as pd
import glob as gb
from fpdf import FPDF
from pathlib import Path

FILEPATHS = gb.glob("invoices/*xlsx*")


def add_invoice_header(pdf, invoice_number, date):
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_number}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)


def add_table_header(pdf, headers, column_withs):
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_fill_color(200, 220, 255)
    
    formatted_headers = [header.replace("_", " ").title() for header in headers]
    for i, header in enumerate(formatted_headers):
        pdf.cell(w=column_withs[i], h=8, txt=header, border=1, fill=True)
    pdf.ln()    


def add_table_rows(pdf, df, column_withs):
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    for _, row in df.iterrows():
        pdf.cell(w=column_withs[0], h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=column_withs[1], h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=column_withs[2], h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=column_withs[3], h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=column_withs[4], h=8, txt=str(row["total_price"]), border=1, ln=1)


def add_total_sum(pdf, df, column_withs):
    total_sum = df["total_price"].sum()

    pdf.cell(w=column_withs[0], h=8, txt="", border=1)
    pdf.cell(w=column_withs[1], h=8, txt="", border=1)
    pdf.cell(w=column_withs[2], h=8, txt="", border=1)
    pdf.cell(w=column_withs[3], h=8, txt="", border=1)
    pdf.cell(w=column_withs[4], h=8, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)    


def add_company_name_logo(pdf):
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)


def create_pdf(filename, invoice_number, date, df):
    pdf = FPDF(orientation="P", unit="mm", format="A4")    
    pdf.add_page()
    
    columns_withs = [30, 70, 32, 30, 30]
    
    add_invoice_header(pdf, invoice_number, date)
        
    add_table_header(pdf, df.columns, columns_withs)
    add_table_rows(pdf, df, columns_withs)    

    add_total_sum(pdf, df, columns_withs)

    add_company_name_logo(pdf)
    
    pdf.output(f"PDFs/{filename}.pdf")


def process_invoices():
    for filepath in FILEPATHS:
        try:
            df = pd.read_excel(filepath, sheet_name="Sheet 1")
            
            filename = Path(filepath).stem
            invoice_number, date = filename.split("-")
            
            create_pdf(filename, invoice_number, date, df)

        except Exception as e:
            print(f"Error processing {filepath}: {e}")


if __name__ == "__main__":
    process_invoices()