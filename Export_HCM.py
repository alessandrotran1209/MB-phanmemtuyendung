import pandas as pd
import os
import base64
import requests
import time
import openpyxl
import csv



while True:
    try:
                a = os.listdir("C:\\MB\\HCM")
                print(a)
                env_delimiter = "YYYYY"  
                delimiter = "XXXXXXXXXX"   
                envr, rest_email = a[0].split(env_delimiter, 1)
                if envr == "DEV":
                        HCM_Import = str("https://prod-32.southeastasia.logic.azure.com:443/workflows/db66c3ce69be4c928841825694b69da1/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QLq3fVRod04ameIBExvyDDcSpuEJKRKBEuGn3-Qv5_U")
                elif envr == "UAT":
                        HCM_Import = str("https://prod-09.southeastasia.logic.azure.com:443/workflows/9b113e43fca34282a64022a8f0a78e63/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aEctYtfIpIsR9_0KL2BQHmb94Z0MVZo-bPFzCmi36BMhcm")
                elif envr == "PROD":
                        HCM_Import = str("https://prod-11.southeastasia.logic.azure.com:443/workflows/10547f16e7de4456ae5c587814c70dff/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=rUoJm7eFdJXnXhS8FBa4RjXqV8noEH3y1vO7nd25WH8")
                owner, rest = rest_email.split(delimiter, 1)
                os.rename(r"C:\\MB\\HCM\\"+a[0],r"C:\\MB\\HCM\\"+rest)
                a[0] = rest[:-4]
                print(a[0])
                owner = owner +"@mbbank.com.vn"
                toxlsx = openpyxl.Workbook()
                ws = toxlsx.active

                with open(f"C:\\MB\\HCM\\{a[0]}.csv",encoding="utf8") as f:
                    reader = csv.reader(f, delimiter=',',)
                    for row in reader:
                        ws.append(row)

                toxlsx.save(f"C:\\MB\\HCM\\{a[0]}.xlsx")
                

                
                # Open the source workbook (A.xlsx)
                source_workbook = openpyxl.load_workbook(f"C:\\MB\\HCM\\{a[0]}.xlsx")
                print('load source')
                source_sheet = source_workbook.active  # Assuming you want to work with the active sheet
                x= 1
                if a[0].startswith('TM') == True:
                    target = "TM"
                    x = 3
                elif a[0].startswith('TL') == True:
                    target = "TL"
                else: target = "TN"


                print('xxx')
                # Open the destination workbook (B.xlsx)
                destination_workbook = openpyxl.load_workbook(f"C:\\MB\\HCM_Template\\{target}.xlsx")
                destination_sheet = destination_workbook.active  # Assuming you want to work with the active sheet

                # Define the starting row index (i)
                i = 2  # Change this to your desired starting row

                # Define the range you want to copy (i to i+10)
                start_row = i
                end_row = i + source_sheet.max_row

                # Define the range where you want to paste (i+5 to i+15)
                paste_start_row = i + x
                paste_end_row = i + destination_sheet.max_row

                # Define the number of columns
                num_columns = source_sheet.max_column  # Assuming both sheets have the same column structure

                # Copy data from source to destination
                for row in range(start_row, end_row + 1):
                    for col in range(1, num_columns + 1):
                        source_cell = source_sheet.cell(row=row, column=col)
                        destination_cell = destination_sheet.cell(row=row + (paste_start_row - start_row), column=col)
                        destination_cell.value = source_cell.value

                # Save the destination workbook
                destination_workbook.save(f"C:\\MB\\HCM\\{a[0]}_done.xlsx")

                # Close both workbooks
                source_workbook.close()
                destination_workbook.close()
                with open(f"C:\\MB\\HCM\\{a[0]}_done.xlsx", "rb") as xlsxbinary1:
                    xlsxfile = base64.b64encode(xlsxbinary1.read()).decode('ascii')
                datasms = {"owner":str(owner),"file":xlsxfile}
                print("Sending to cloud")
                requests.post(HCM_Import, json = datasms)
                os.remove(f"C:\\MB\\HCM\\{a[0]}_done.xlsx")
                os.remove(f"C:\\MB\\HCM\\{a[0]}.xlsx")
                os.remove(f"C:\\MB\\HCM\\{a[0]}.csv")

                print("Done")
    except Exception as error:
        print(error)
        print(str(time.ctime()))
        time.sleep(1)
        pass

            


