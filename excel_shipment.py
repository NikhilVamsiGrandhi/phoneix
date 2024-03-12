import fitz
import openpyxl

def value_from_pdf(file_path):
    doc = fitz.open(rf"{file_path}")
    page = doc[-1]
    text = page.get_text()
    list = text.split('\n')
    value = list[-2]
    value = value.replace('-','').replace(',','')
    value = int(round(float(value)))
    return value

def sheet_shipment(excel_path,pdf_path):
    print(pdf_path)
    workbook = openpyxl.load_workbook(excel_path)

    sheet_names = workbook.sheetnames
    last_sheet_name = sheet_names[-1]
    worksheet = workbook[last_sheet_name]

    worksheet['c44']= value_from_pdf(pdf_path)
    value = worksheet['c44'].value
    print(value)
    workbook.save(excel_path)

# sheet_shipment(r"C:\Users\NikhilVamsiGrandhi\OneDrive - Kanerika Software\Desktop\phoenix\excel_files\DailyIntake.xlsx",r"C:\Users\NikhilVamsiGrandhi\OneDrive - Kanerika Software\Desktop\phoenix\pdfs\Shipment_Product_Purchase_By_Date[8_Mar].pdf")


