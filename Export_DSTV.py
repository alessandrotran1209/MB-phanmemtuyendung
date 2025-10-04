import pandas as pd
import os
import base64
import requests
import time

exportdstv_url = str("https://prod-41.southeastasia.logic.azure.com:443/workflows/d38dd543dfc549d8a7d0b31c5b4ab5b8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=YGv9-kvy1a69XLRexkKSVlHWLUw2nKrImgHpPJrLPUE")
while True:
    try:
                a = os.listdir("C:\\MB\\Export_DSTV")
                print(a)
                env_delimiter = "YYYYY"  
                delimiter = "XXXXXXXXXX"   
                envr, rest_email = a[0].split(env_delimiter, 1)
                if envr == "DEV":
                        exportdstv_url = str("https://prod-41.southeastasia.logic.azure.com:443/workflows/d38dd543dfc549d8a7d0b31c5b4ab5b8/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=YGv9-kvy1a69XLRexkKSVlHWLUw2nKrImgHpPJrLPUE")
                elif envr == "UAT":
                        exportdstv_url = str("https://prod-27.southeastasia.logic.azure.com:443/workflows/f2da64c320464ed0ba9cbb5b5c22b176/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=f6lUGyDq-IqsA-BzNJI6wE9DGrY_hDj75PgiMJCAdkU")
                elif envr == "PROD":
                        exportdstv_url = str("https://prod-57.southeastasia.logic.azure.com:443/workflows/a8d1bdc8dbc34f498d7dc0cab6d97f3a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=lr37rUL7BvFR8eMLcoTZ0zTb-eKO4tRPf7YEw3VpaZI")
                owner, rest = rest_email.split(delimiter, 1)
                os.rename(r"C:\\MB\\Export_DSTV\\"+a[0],r"C:\\MB\\Export_DSTV\\"+rest)
                a[0] = rest[:-4]
                print(a[0])
                owner = owner +"@mbbank.com.vn"
                
                read_file_product = pd.read_csv (r"C:\\MB\\Export_DSTV\\"+a[0],encoding='iso-8859-1')
                newname = a[0][:-4]
                print(newname)
                read_file_product.to_excel (f'C:\\MB\\Export_DSTV\\{newname}.xlsx', index = None, header=True)
                with open(f'C:\\MB\\Export_DSTV\\{newname}.xlsx', "rb") as xlsxbinary:
                    xlsxfile = base64.b64encode(xlsxbinary.read()).decode('ascii')
                os.remove(f"C:\\MB\\Export_DSTV\\{newname}.xlsx")
                os.remove(r"C:\\MB\\Export_DSTV\\"+a[0])
                datasms = {"owner":str(newname),"file":xlsxfile}
                print("Sending to Dataverse")
                requests.post(exportdstv_url, json = datasms)
                print("Done")
    except Exception as error:
        print(error)
        print(str(time.ctime()))
        time.sleep(1)
        pass

            