import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
from time import strftime


filepaths = glob.glob("excel_invoces/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]

    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.: {invoice_nr}", ln=1)

    current_date = strftime("%b %d, %Y")
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {current_date}")
    pdf.ln(20)

    db = pd.read_excel(filepath, sheet_name="Sheet 1")
    headers = list(db.columns)
    titles = [item.replace("_", " ").title() for item in headers]
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt=titles[0], border=1)
    pdf.cell(w=70, h=8, txt=titles[1], border=1)
    pdf.cell(w=30, h=8, txt=titles[2], border=1)
    pdf.cell(w=30, h=8, txt=titles[3], border=1)
    pdf.cell(w=30, h=8, txt=titles[4], border=1, ln=1)

    for index, row in db.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = db["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1)
    pdf.ln(20)

    pdf.set_font(family="Times", size=20, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=f"Total Price {total_sum}")
    pdf.ln(10)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=36, h=8, txt=f"Super Python")
    pdf.image("img/python_logo.png", w=10)

    pdf.output(f"pdfs/{filename}.pdf")
