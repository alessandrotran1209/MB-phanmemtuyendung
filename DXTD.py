import pandas as pd
import time
import os
import shutil
import requests
from datetime import datetime

envr = "dev"

null = ''


while True:
    try:
        try:
#Preprocess--------------------------------------------------------
                a = os.listdir("C:\\MB\\DXTD")
                if len(a) > 0:
                        blank = False
                else: blank = True
                print(a)
                env_delimiter = "YYYYY"  
                delimiter = "XXXXXXXXXX"   
                envr, rest_email = a[0].split(env_delimiter, 1)
                if envr == "DEV":
                        dxtd = str("https://prod-60.southeastasia.logic.azure.com:443/workflows/431c45cdfde44a47aada40a0c5e7a286/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Eer5J9IXl2ZEU-TO-mjjK9R_2iw33ajJjFPUhsPnYBg") 
                elif envr == "UAT":
                        dxtd = str("https://prod-32.southeastasia.logic.azure.com:443/workflows/45f539f041f34f269740df1ebc6624af/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_D18EOx3pBZlSfWFpRMk94xtwyL484KZOpiBTtOkuZs")
                elif envr == "PROD":
                        dxtd = str("https://prod-41.southeastasia.logic.azure.com:443/workflows/a32e31c6f2bb49c3a7b329d149d4292a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Rqkaj6BKp4X7sbFzYyNgrILAeZewjbwyO4zcELMa_1o")
                                
                owner, rest = rest_email.split(delimiter, 1)

                os.rename(r"C:\\MB\\DXTD\\"+a[0],r"C:\\MB\\DXTD\\"+rest)
                a[0] = rest
                owner = owner +"@mbbank.com.vn"

                def vld1(inputtime): #time validate
                        a = str(inputtime)
                        if a == None or a==null or a =="None" or a == "NaN" or a == "nan":
                                a = '01/01/1900'
                        if len(str(a)) == 4:
                                a = '01/01/'+str(a)
                        elif len(str(a)) == 6:
                                a = '01/0'+str(a)
                        elif len(str(a)) == 7:
                                a = '01/'+str(a)
                        elif "-" in str(a):
                                a= str(a)
                                a = datetime.strptime(a, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
                        elif "." in str(a):
                                a=str(a).replace(".","/")
                        else: a = a
                        return str(a)
                df = pd.read_excel(r"C:\\MB\\DXTD\\"+a[0])
                for index, row in df.iterrows():
                        data = { "owner": str(owner),
                        "CV_layout": "importdxtd",
                "username": "admin",
                        "password": "@MB2023",
                "MaDeXuat": str(row.iloc[1]),
                "LyDoDeXuat": str(row.iloc[2]),
                "NgayDeXuat": vld1(row.iloc[3]),
                "DonViQuanLy": str(row.iloc[4]),
                "DVTT_1": str(row.iloc[5]),
                "DVTT_2": str(row.iloc[6]),
                "DVTT_3": str(row.iloc[7]),
                "DVTT_4": str(row.iloc[8]),
                "Job": str(row.iloc[9]),
                "SoLuongPheDuyet": str(row.iloc[13]),
                "CanBoDeXuat": str(row.iloc[25]),
                "CanBoPheDuyet": str(row.iloc[26]),
                "NgayPheDuyet": vld1(row.iloc[22]),
                "DauMoiTuyenDung": str(row.iloc[10])
        }
                        print(f'Sending proposal data #{index+1} to Dataverse')
                        requests.post(dxtd, json = data)
                        print('DONE')
                os.remove(r"C:\\MB\\DXTD\\"+a[0])   
                raise ValueError('Batch processing finished')

        except Exception as error1:
            if blank == False:
                    os.remove(r"C:\\MB\\DXTD\\"+a[0])
                    erroritem = {"filename": str(a[0]), "owner": str(owner)}
                    print('Moved to error folder')
                    print(error1)
                    # requests.post(errorurl, json = erroritem)  
            else:
                    raise ValueError(str(datetime.now()))

    except Exception as error:
        print(error)
        print("Please upload DXTD file")
        time.sleep(2)
        pass


