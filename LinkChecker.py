import openpyxl
import requests
from openpyxl.styles import Font

#def user_input():
    #path = input("Enter the full directory path to the Excel file we're testing: ").strip()
    #file_name = input("Enter the file name of Excel file we're testing: ").strip()
    #test_hyperlinks(r"{path}+{file_name}")

def test_hyperlinks(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    rowCount=1

    for row in(ws.iter_rows(min_row=2, max_col=2, values_only=True)):
        rowCount+=1
        cell_value = row[1]
        if cell_value:
            try:
                response=requests.head(cell_value, allow_redirects=True)
                #status_code=requests.Response
                status_code=response.status_code
                

                
                if 200<=status_code<400:
                    result="Active"
                    ws.cell(row=rowCount, column=3).value=result
                    ws.cell(row=rowCount, column=3).font=Font(color=r"00FF00")
                    #ws.cell(row=rowCount, column=3, value=result).font=Font(color="00FF00")

                else:
                    result=f"Inactive ({status_code})"
                    ws.cell(row=rowCount, column=3).value=result
                    ws.cell(row=rowCount, column=3).font=Font(color=r"FF0000")

                    #ws.cell(row=rowCount, column=3, value=result).font=Font(color="FF0000")
                    

            except:
                    requests.RequestException
                    result=f"Error"
                    ws.cell(row=rowCount, column=3).value=result
                    ws.cell(row=rowCount, column=3).font=Font(color=r"FFA500")

                    #ws.cell(row=row[0], column=3, value=result).font= Font(color="FFA500")

            
    wb.save(file_path)

test_hyperlinks(r"C:\Users\Flip\Desktop\LinkChecker\TestLinks.xlsx")

#This function not currently working
#user_input()