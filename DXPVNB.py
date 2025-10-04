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
                a = os.listdir("C:\\MB\\DXPVNB")
                if len(a) > 0:
                        blank = False
                else: blank = True
                print(a)
                env_delimiter = "YYYYY"  
                delimiter = "XXXXXXXXXX"   
                envr, rest_email = a[0].split(env_delimiter, 1)
                if envr == "DEV":
                        dxpvnb = str("https://prod-21.southeastasia.logic.azure.com:443/workflows/54dd1e6e7a60413a9f2bce039d0ee4a0/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nDiEM_rY8FtQHyV5VXkkRfhLjDitG8cEQbU0OpJyYHg")
                elif envr == "UAT":
                        dxpvnb = str("https://prod-47.southeastasia.logic.azure.com:443/workflows/1199d1568b49460887abca3bb04e29eb/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=onf0GjEYII8H1btWC6DZSjxTLI0Jo7C5R90Uvz5d3ZA") 
                elif envr == "PROD":
                        dxpvnb = str("https://prod-14.southeastasia.logic.azure.com:443/workflows/c57fa74409cb48f28a099c8c322d4fcc/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EHYIdfOLKalxuzH23cRa7Sw1Gk56eGqgdGWVUEu27Kw") 
                
                owner, rest = rest_email.split(delimiter, 1)

                os.rename(r"C:\\MB\\DXPVNB\\"+a[0],r"C:\\MB\\DXPVNB\\"+rest)
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
                df = pd.read_excel(r"C:\\MB\\DXPVNB\\"+a[0])
                for index, row in df.iterrows():
                        data = { "owner": str(owner),
                        "CV_layout": "importdxpv",
                "username": "admin",
                        "password": "@MB2023",
                        "MaDeXuatPhongVan": str(row.iloc[1]),
                "TenNV": str(row.iloc[3]),
                "MaNV": str(row.iloc[2]),
                "CCCD": str(row.iloc[2]),
                "HT_ChucDanh": str(row.iloc[14]),
                "HT_DonViQuanLy": str(row.iloc[13]),
                "HT_ViTri": str(row.iloc[15]),
                "DXM_ChucDanh": str(row.iloc[17]),
                "DXM_DonViQuanLy": str(row.iloc[16]),
                "DXM_ViTri": str(row.iloc[18]),
                "NgayVao_MB": vld1(row.iloc[12]),
                "HinhThuc": str(row.iloc[10]),
                "YKien_PTNL": str(row.iloc[19]),
                "NgayDeXuat": vld1(row.iloc[4]), 
                "NgaySinh": vld1(row.iloc[11]) }
                        print(f'Sending data of candidate #{index+1} to Dataverse')
                        for key, value in data.items():
                                key_type = type(key).__name__
                                value_type = type(value).__name__
                                print(f"Key: {key},Value: {value}, Type: {value_type}")
                        requests.post(dxpvnb, json = data)
                        print('DONE')
                os.remove(r"C:\\MB\\DXPVNB\\"+a[0])                    
                raise ValueError('Batch processing finished')

        except Exception as error1:
            if blank == False:
                    os.remove(r"C:\\MB\\DXPVNB\\"+a[0])
                    erroritem = {"filename": str(a[0]), "owner": str(owner)}
                    print('Moved to error folder')
                    print(error1)
                    # requests.post(errorurl, json = erroritem)  
            else:
                    raise ValueError(str(datetime.now()))

    except Exception as error:
        print(error)
        print("Please upload DXPVNB file")
        time.sleep(2)
        pass


