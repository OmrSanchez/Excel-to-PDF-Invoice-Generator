import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
	pdf = FPDF(orientation="P", unit="mm", format="A4")
	pdf.add_page()

	filename = Path(filepath).stem
	invoice_nbr, date = filename.split("-")

	pdf.set_font(family="Times", size=16, style="B")
	pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nbr}",  align="L", ln=1)

	pdf.set_font(family="Times", size=16, style="B")
	pdf.cell(w=50, h=8, txt=f"Date: {date} ", align="L", ln=5)

	df = pd.read_excel(filepath, sheet_name="Sheet 1")

	headers = df.columns
	headers = [item.replace("_", " ").capitalize() for item in headers]

	pdf.set_font(family="Times", size=10, style="B")
	pdf.set_text_color(80, 80, 80)
	pdf.cell(w=30, h=8, txt=headers[0], border=1)
	pdf.cell(w=50, h=8, txt=headers[1], border=1)
	pdf.cell(w=40, h=8, txt=headers[2], border=1, align="R")
	pdf.cell(w=30, h=8, txt=headers[3], border=1, align="R")
	pdf.cell(w=30, h=8, txt=headers[4], ln=1, border=1, align="R")


	for index, row in df.iterrows():
		pdf.set_font(family="Times", size=10)
		pdf.set_text_color(80, 80, 80)
		pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
		pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
		pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1, align="R")
		pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1, align="R")
		pdf.cell(w=30, h=8, txt=str(row["total_price"]), ln=1, border=1, align="R")

	pdf.output(f"PDFs/{filename}.pdf")


