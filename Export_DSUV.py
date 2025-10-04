import pandas as pd
import os
import base64
import requests
import time


while True:
    try:
                a = os.listdir("C:\\MB\\DSUV")
                print(a)
                env_delimiter = "YYYYY"  
                delimiter = "XXXXXXXXXX"   
                envr, rest_email = a[0].split(env_delimiter, 1)
                if envr == "DEV":
                        dsuv_url = str("https://prod-01.southeastasia.logic.azure.com:443/workflows/2aebd88355134fe0a569be1354312b71/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=b6PGwA7Vm3iVGi2qvfnb80JDgwGbrSwEzKsk2G5Psdc")
                elif envr == "UAT":
                        dsuv_url = str("https://prod-35.southeastasia.logic.azure.com:443/workflows/84d2ca00b33040429f36a560d3ef184e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=bJNqRmz8KGyflCiEEaY0k6JBPkX2rba7ULgHPATk-X0")
                elif envr == "PROD":
                        dsuv_url = str("https://prod-54.southeastasia.logic.azure.com:443/workflows/4540882be7f24f818e0c9465566698e4/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-lm77R66qe4yjyBj0G1KXwOTrvBviIM_RVVRjs4stH4")
                owner, rest = rest_email.split(delimiter, 1)
                os.rename(r"C:\\MB\\THKQTD\\"+a[0],r"C:\\MB\\THKQTD\\"+rest)
                a[0] = rest[:-4]
                print(a[0])
                read_file_product = pd.read_csv (r"C:\\MB\\DSUV\\"+a[0],encoding='iso-8859-1')
                newname = a[0][:-4]
                print(newname)
                read_file_product.to_excel (f'C:\\MB\\DSUV\\{newname}.xlsx', index = None, header=True)
                with open(f'C:\\MB\\DSUV\\{newname}.xlsx', "rb") as xlsxbinary:
                    xlsxfile = base64.b64encode(xlsxbinary.read()).decode('ascii')
                os.remove(f"C:\\MB\\DSUV\\{newname}.xlsx")
                os.remove(r"C:\\MB\\DSUV\\"+a[0])
                datasms = {"owner":str(newname),"file":xlsxfile}
                print("Sending to Dataverse")
                requests.post(dsuv_url, json = datasms)
                print("Done")
    except Exception as error:
        print(error)
        print(str(time.ctime()))
        time.sleep(1)
        pass

            


