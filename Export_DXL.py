import pandas as pd
import os
import base64
import requests
import time

exportExport_DXL_url = str("https://prod-52.southeastasia.logic.azure.com:443/workflows/6624043dfb46423da91b7842e30255bf/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=V3BXSjb79EW8Ab-NjbL-G_bY0a3oU3pA0wBPGdAjVrE")

while True:
    try:
                a = os.listdir("C:\\MB\\Export_DXL")
                read_file_product = pd.read_csv (r"C:\\MB\\Export_DXL\\"+a[0],encoding='utf8')
                newname = a[0][:-4]
                print(newname)
                read_file_product.to_excel (f'C:\\MB\\Export_DXL\\{newname}.xlsx', index = None, header=True)
                with open(f'C:\\MB\\Export_DXL\\{newname}.xlsx', "rb") as xlsxbinary:
                    xlsxfile = base64.b64encode(xlsxbinary.read()).decode('ascii')
                os.remove(f"C:\\MB\\Export_DXL\\{newname}.xlsx")
                os.remove(r"C:\\MB\\Export_DXL\\"+a[0])
                dataExport_DXL = {"owner":str("pwappadmin"),"file":xlsxfile}
                print("Sending to Dataverse")
                requests.post(exportExport_DXL_url, json = dataExport_DXL)
                print("Done")
    except Exception as error:
        print(error)
        print(str(time.ctime()))
        time.sleep(1)
        pass