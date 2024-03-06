import openpyxl
from pdf import value_from_pdf
from datetime import datetime
from openpyxl.styles import NamedStyle

today = datetime.now().day
month = datetime.now().strftime("%B")
month = month[0:3]
year = datetime.now().year

dict_day = {
    "TRANSPORT":"j26",
    "1% CASH HANDLING FEE": "j38",
    "Fuel Adjustment Factor":"j41",
}
dict_month = {
    "TRANSPORT":"j27",
    "1% SORTING FEE":"j30",
    "1% CASH HANDLING FEE": "j39",
    "Fuel Adjustment Factor":"j42",
    "N":"c30",
    "B":"c34"
}

def value_update(sheet,timeFrame,keyword,pdf_path,total=False):
    keyword_value = value_from_pdf(keyword,pdf_path,total)
    keyword_col = dict_day[keyword] if timeFrame=="day" else dict_month[keyword]
    sheet[keyword_col] = keyword_value
    

def new_excel_sheet(excel_path,pdf_path ,timeFrame):
    workbook = openpyxl.load_workbook(excel_path)

    sheet_names = workbook.sheetnames
    last_sheet_name = sheet_names[-1]
    last_sheet = workbook[last_sheet_name]

    new_sheet = workbook.copy_worksheet(last_sheet)
    new_sheet.title = "DailyIntake"+str(today)+month+str(year)

    # new_sheet['b1'].value=datetime(year,datetime.now().month,today)
    # new_sheet['b1'].number_format = 'dd-mm-yyyy'
    # prev_wrkdys = new_sheet['s3'].value
    # new_sheet['s3']=(prev_wrkdys+1)
    
    # value_update(new_sheet,timeFrame,"TRANSPORT",pdf_path)
    # value_update(new_sheet,timeFrame,"1% CASH HANDLING FEE",pdf_path)
    # value_update(new_sheet,timeFrame,"Fuel Adjustment Factor",pdf_path)
    # if timeFrame!="day":
    #     value_update(new_sheet,timeFrame,"1% SORTING FEE",pdf_path)
    #     value_update(new_sheet,timeFrame,"N",pdf_path,total=True)
    #     value_update(new_sheet,timeFrame,"B",pdf_path,total=True)
    
    workbook.save(excel_path)

# new_excel_sheet(r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\excel_files\DailyIntake_test.xlsx",r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\pdfs\Daily_Product_Purchase_By_Date[2Mar].pdf","day")
excel_path = r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\excel_files\DailyIntake.xlsx"
pdf_path = r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\pdfs\Daily_Product_Purchase_By_Date[2Mar].pdf"
timeFrame = "day"
new_excel_sheet(excel_path, pdf_path, timeFrame)


# import openpyxl
# from datetime import datetime
# from openpyxl.styles import NamedStyle

# def duplicate_sheet(original_sheet, new_sheet):
#     for row in original_sheet.iter_rows(min_row=1, max_row=original_sheet.max_row, values_only=True):
#         new_sheet.append(row)

#     for col in original_sheet.columns:
#         new_sheet.column_dimensions[col[0].column_letter].width = original_sheet.column_dimensions[col[0].column_letter].width

#     for row in original_sheet.iter_rows(min_row=1, max_row=original_sheet.max_row, min_col=1, max_col=original_sheet.max_column):
#         for cell in row:
#             new_cell = new_sheet[cell.coordinate]

#             # Copy cell value
#             new_cell.value = cell.value

#             # Copy style
#             new_style = NamedStyle(name="duplicate_style")
#             new_style.font = cell.font
#             new_style.border = cell.border
#             new_style.fill = cell.fill
#             new_style.number_format = cell.number_format
#             new_style.alignment = cell.alignment

#             new_sheet[cell.coordinate].style = new_style

#     for col_num, col in enumerate(original_sheet.columns, 1):
#         for cell in col:
#             if isinstance(cell.value, (int, float)):
#                 continue
#             if cell.number_format:
#                 new_sheet.cell(row=cell.row, column=col_num).number_format = cell.number_format


# def new_excel_sheet(excel_path, pdf_path, timeFrame):
#     workbook = openpyxl.load_workbook(excel_path)

#     # Assuming 'today', 'month', and 'year' are defined
#     today = datetime.now().day
#     month = datetime.now().month
#     year = datetime.now().year

#     sheet_names = workbook.sheetnames
#     last_sheet_name = sheet_names[-1]
#     last_sheet = workbook[last_sheet_name]

#     new_sheet = workbook.copy_worksheet(last_sheet)
#     new_sheet.title = "DailyIntake" + str(today) + str(month) + str(year)

#     new_sheet['B1'] = str(today) + "-" + str(month) + "-" + str(year)
#     prev_wrkdys = new_sheet['S3'].value  # Assuming the column 'S' represents days
#     new_sheet['S3'] = (prev_wrkdys + 1)

#     duplicate_sheet(last_sheet, new_sheet)

#     workbook.save(excel_path)

# # Example usage:
# excel_path = r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\excel_files\DailyIntake_test.xlsx"
# pdf_path = r"C:\Users\KSPL\Desktop\automation\Python\ScrapIT_PTG\pdfs\Daily_Product_Purchase_By_Date[2Mar].pdf"
# timeFrame = "day"
# new_excel_sheet(excel_path, pdf_path, timeFrame)
