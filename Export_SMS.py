import pandas as pd
import os
import base64
import requests
import time

exportsms_url = str("https://prod-18.southeastasia.logic.azure.com:443/workflows/83574bda358c4e999f94eb88be204dcb/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IkhWVt85XQB97jU_LDgD-aKtteaMMEGa-Z8_rM1blpI")

while True:
    try:
                a = os.listdir("C:\\MB\\SMS")
                read_file_product = pd.read_csv (r"C:\\MB\\SMS\\"+a[0],encoding='utf8',dtype=str)
                newname = a[0][:-4]
                print(newname)
                read_file_product.to_excel (f'C:\\MB\\SMS\\{newname}.xlsx', index = None, header=True)
                with open(f'C:\\MB\\SMS\\{newname}.xlsx', "rb") as xlsxbinary:
                    xlsxfile = base64.b64encode(xlsxbinary.read()).decode('ascii')
                os.remove(f"C:\\MB\\SMS\\{newname}.xlsx")
                os.remove(r"C:\\MB\\SMS\\"+a[0])
                datasms = {"owner":str("pwappadmin"),"file":xlsxfile}
                print("Sending to Dataverse")
                requests.post(exportsms_url, json = datasms)
                print("Done")
    except Exception as error:
        print(error)
        print(str(time.ctime()))
        time.sleep(1)
        pass

            


