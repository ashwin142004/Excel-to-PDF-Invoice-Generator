import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('Invoices/*.xlsx')

for filepath in filepaths: 
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    filename = Path(filepath).stem
    
    invoice_no, date = filename.split('-')
    
    pdf.set_font('Times', 'B', 16)
    pdf.cell(50, 8, f'Invoice No. {invoice_no}' , ln = 1) 

    pdf.set_font('Times', 'B', 16)
    pdf.cell(50, 8, f'Date {date}', ln = 1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    columns = df.columns
    columns = [item.replace('_', ' ').title() for item in columns]
    pdf.set_font('Times', 'B', 12)

    pdf.cell(25, 8, f'{columns[0]}' , border=1)
    pdf.cell(50, 8, f'{columns[1]}', border=1)
    pdf.cell(40, 8, f'{columns[2]}', border=1)
    pdf.cell(30, 8, f'{columns[3]}', border=1)
    pdf.cell(30, 8, f'{columns[4]}', ln = 1, border=1)

    for index, row in df.iterrows():
        pdf.set_font('Times', '', 12)
        pdf.cell(25, 8, f'{row["product_id"]}' , border=1)
        pdf.cell(50, 8, f'{row["product_name"]}', border=1)
        pdf.cell(40, 8, f'{row["amount_purchased"]}', border=1)
        pdf.cell(30, 8, f'{row["price_per_unit"]}', border=1)
        pdf.cell(30, 8, f'{row["total_price"]}', ln = 1, border=1)
    
    total_sum = df['total_price'].sum()
    pdf.set_font('Times', '', 12)
    pdf.cell(25, 8, "" , border=1)
    pdf.cell(50, 8, "", border=1)
    pdf.cell(40, 8, "", border=1)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(30, 8, str(total_sum), ln = 1, border=1)

    pdf.set_font('Times', 'B', 10)
    pdf.cell(30, 8, f"The total price is {total_sum}", ln = 1)

    pdf.set_font('Times', 'B', 10)
    pdf.cell(30, 8, f"Thank you for your purchase", ln = 1)


    pdf.output(f'PDFs/{filename}.pdf')
