import PyPDF2
from openpyxl import Workbook
import os


workbook_stock = Workbook()
workbook_tax = Workbook()

# Select the active sheet
sheet_stock = workbook_stock.active
sheet_tax = workbook_tax.active

stock_feature=["Date","Order Number","Trade Number","Security/Contract Description","Buy/Sell","Qty"]
tax_feature=['Date','STT/CT','Brokerage','Transaction and Clearing Charges','Stamp Duty','Sebi Fee/RM','Taxable value of supply','Cgst@9%','Sgst@9%','Net Amount']

sheet_tax.append(tax_feature)
sheet_stock.append(stock_feature)
folder_path = r'C:\Users\Dharavat hanumanth\Desktop\saiteja_app\stocks'
for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'): # Check if the file is a PDF file
        file_path = os.path.join(folder_path, filename) # Get the full file path
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)

            page = pdf_reader.pages[0]

            text = page.extract_text()

            lines = text.split('\n')
            taxes_all=lines[-18:-9]
            date=lines[-9].split()[2]

            stocks_all =lines[57:-21]
            for i in stocks_all:
                stock_temp=i.split()
                stock=[date,stock_temp[0],stock_temp[2],stock_temp[4]+stock_temp[5]+stock_temp[6]+stock_temp[7],stock_temp[9],stock_temp[10]]
                #print(stock)
                #['1300000006444723', '78169183', 'StateBankofIndia', 'Buy', '21'] output stock
                sheet_stock.append(stock)
            tax_temp=[date]
            for i in taxes_all:
                tax_temp.append(i.split()[-1])

            sheet_tax.append(tax_temp)

            # Save the workbook of stock
workbook_stock.save('stock.xlsx')
workbook_tax.save('tax.xlsx')