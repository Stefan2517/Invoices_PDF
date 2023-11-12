# Am instalat openpyxl
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Sunt in acelasi folder cu main.py, altfel precizam calea
filepaths = glob.glob("*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    # invoices_nr = filename.split("-")[0]
    # Astfel date i-a indexul 1, iar codul e mai scurt
    invoices_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoices_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    # Le-am transformat in lista
    #columns_var = list(df.columns)

    columns_var = df.columns
    columns_var = [item.replace("-"," ").title() for item in columns_var]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=8, txt=columns_var[0], border=1)
    pdf.cell(w=70, h=8, txt=columns_var[1], border=1)
    pdf.cell(w=35, h=8, txt=columns_var[2], border=1)
    pdf.cell(w=30, h=8, txt=columns_var[3], border=1)
    pdf.cell(w=30, h=8, txt=columns_var[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        # AttributeError: 'int' object has no attribute 'replace' in caz ca nu pun str inainte de row
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
