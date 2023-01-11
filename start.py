import gspread
import os
import sys
from docxtpl import DocxTemplate

#Change working directory from default to current run path
os.chdir(sys.path[0])

#Open spreadsheet function
def openSpreadSheet(cerdFileName, sheetName, workSheetName) :
    account = gspread.service_account(cerdFileName)
    sheet = account.open(sheetName)
    worksheet = sheet.worksheet(workSheetName)
    return worksheet

#Main run command
def run() :
    #Open Template
    doc = DocxTemplate('.\\format\\format.docx')

    fileName = input("Insert file name name : ").strip()
    sheetName = input("Insert  sheet name : ").strip()

    #Open spreadsheet
    worksheet = openSpreadSheet("cerds.json", fileName, sheetName) 
    cell = worksheet.find('Status')
    status_row = cell.row
    status_col = cell.col
    records = worksheet.get_all_records()
    
    #Create docx file
    for n,data in enumerate(records) :
        if data['Status'] == '' :
            doc.render(data)
            date, time = data['Timestamp'].split(' ')
            date = date.replace("/", ".")
            time = time.replace(":", ";")
            try :
                os.mkdir(f'.\\output\\{date}')
            except Exception as e :
                if "WinError 183" in str(e) :
                    print("Already has folder: passing...")
                else :
                    print(e)

            doc.save(f'.\\output\\{date}\\{time}.docx')

            worksheet.update_cell(n+status_row+1, status_col, 'Downloaded')

if __name__ == "__main__" :
    run()
