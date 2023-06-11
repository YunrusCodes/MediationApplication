import os
import PyPDF2
from docx2pdf import convert  # pip install docx2pdf
import win32com.client  # pip install pywin32


def excel_to_pdf(excel_path, pdf_path):
    print(excel_path)
    print(pdf_path)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.Workbooks.Open(excel_path)
    wb = excel.ActiveWorkbook
    ws = wb.ActiveSheet
    ws.ExportAsFixedFormat(0, pdf_path)
    wb.Close()
    excel.Quit()

import sys

pdfs = []
# 檢查是否傳遞了引數
if len(sys.argv) > 1:
    # 遍例所有引數（從索引位置1開始）
    for arg in sys.argv[1:]:
        if arg.endswith('docx'):
            pdf_path = arg.replace('.docx','.pdf')
            convert(arg,pdf_path)
            pdfs.append(pdf_path)
        elif arg.endswith('xlsx'):
            pdf_path = arg.replace('.xlsx','.pdf')
            excel_to_pdf(arg,pdf_path)
            pdfs.append(pdf_path)
        else:
            print(f"無效的引數:{arg}")
else:
    print("未傳遞引數")

# 合併PDF檔案
pdf_writer = PyPDF2.PdfWriter()
for pdf_file in pdfs:
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

# 刪除原始的PDF檔案
for file_to_remove in pdfs:
    os.remove(file_to_remove)

# 寫入合併後的PDF檔案
merged_pdf_path = pdfs[0]
with open(merged_pdf_path, 'wb') as merged_pdf_file:
    pdf_writer.write(merged_pdf_file)
