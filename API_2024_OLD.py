from fastapi import FastAPI, HTTPException, Depends, Request
from fastapi.responses import JSONResponse
import uvicorn
import base64
import csv
import io
from io import BytesIO
import re
import zipfile
from PIL import Image
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import psutil
from openpyxl import load_workbook
import win32com.client as win32
import base64
from PIL import ImageGrab
import pythoncom
from xlsx2html import xlsx2html
import pdfkit
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
import math


def convert_to_ascii(vietnamese_string):
    # Define a dictionary mapping Vietnamese characters to their ASCII equivalents
    vietnamese_ascii_mapping = {
        'á': 'a', 'à': 'a', 'ả': 'a', 'ã': 'a', 'ạ': 'a',
        'ă': 'a', 'ắ': 'a', 'ằ': 'a', 'ẳ': 'a', 'ẵ': 'a', 'ặ': 'a',
        'â': 'a', 'ấ': 'a', 'ầ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ậ': 'a',
        'đ': 'd',
        'é': 'e', 'è': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ẹ': 'e',
        'ê': 'e', 'ế': 'e', 'ề': 'e', 'ể': 'e', 'ễ': 'e', 'ệ': 'e',
        'í': 'i', 'ì': 'i', 'ỉ': 'i', 'ĩ': 'i', 'ị': 'i',
        'ó': 'o', 'ò': 'o', 'ỏ': 'o', 'õ': 'o', 'ọ': 'o',
        'ô': 'o', 'ố': 'o', 'ồ': 'o', 'ổ': 'o', 'ỗ': 'o', 'ộ': 'o',
        'ơ': 'o', 'ớ': 'o', 'ờ': 'o', 'ở': 'o', 'ỡ': 'o', 'ợ': 'o',
        'ú': 'u', 'ù': 'u', 'ủ': 'u', 'ũ': 'u', 'ụ': 'u',
        'ư': 'u', 'ứ': 'u', 'ừ': 'u', 'ử': 'u', 'ữ': 'u', 'ự': 'u',
        'ý': 'y', 'ỳ': 'y', 'ỷ': 'y', 'ỹ': 'y', 'ỵ': 'y',
        'Á': 'A', 'À': 'A', 'Ả': 'A', 'Ã': 'A', 'Ạ': 'A',
        'Ă': 'A', 'Ắ': 'A', 'Ằ': 'A', 'Ẳ': 'A', 'Ẵ': 'A', 'Ặ': 'A',
        'Â': 'A', 'Ấ': 'A', 'Ầ': 'A', 'Ẩ': 'A', 'Ẫ': 'A', 'Ậ': 'A',
        'Đ': 'D',
        'É': 'E', 'È': 'E', 'Ẻ': 'E', 'Ẽ': 'E', 'Ẹ': 'E',
        'Ê': 'E', 'Ế': 'E', 'Ề': 'E', 'Ể': 'E', 'Ễ': 'E', 'Ệ': 'E',
        'Í': 'I', 'Ì': 'I', 'Ỉ': 'I', 'Ĩ': 'I', 'Ị': 'I',
        'Ó': 'O', 'Ò': 'O', 'Ỏ': 'O', 'Õ': 'O', 'Ọ': 'O',
        'Ô': 'O', 'Ố': 'O', 'Ồ': 'O', 'Ổ': 'O', 'Ỗ': 'O', 'Ộ': 'O',
        'Ơ': 'O', 'Ớ': 'O', 'Ờ': 'O', 'Ở': 'O', 'Ỡ': 'O', 'Ợ': 'O',
        'Ú': 'U', 'Ù': 'U', 'Ủ': 'U', 'Ũ': 'U', 'Ụ': 'U',
        'Ư': 'U', 'Ứ': 'U', 'Ừ': 'U', 'Ử': 'U', 'Ữ': 'U', 'Ự': 'U',
        'Ý': 'Y', 'Ỳ': 'Y', 'Ỷ': 'Y', 'Ỹ': 'Y', 'Ỵ': 'Y',
    }
    
    # Initialize an empty string to store the converted string
    ascii_string = ''
    
    # Loop through each character in the Vietnamese string
    for char in vietnamese_string:
        # If the character is found in the mapping, replace it with its ASCII equivalent
        if char in vietnamese_ascii_mapping:
            ascii_string += vietnamese_ascii_mapping[char]
        # If the character is not found in the mapping, keep it unchanged
        else:
            ascii_string += char
    
    return ascii_string


def kill():
      for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                  proc.kill()

def null_vld(value1):
    if value1 is None:
        return ""
    elif isinstance(value1, str):
        return value1
    elif isinstance(value1, (int, float)):
        return str(int(value1))
    else:
        return "unknown"

def vld(mode,date_string):
    try:    
        if mode == 1:
                return  datetime.strptime(str(date_string), "%m/%Y").isoformat()
        elif mode == 2:
                return  datetime.strptime(str(date_string), "%Y").isoformat()
        else:
                return datetime.strptime(str(date_string), "%d/%m/%Y").isoformat()
    except:
          return "1900-01-01T00:00:00" 
null = ""

def csv2xlsx(CSVstring,sms:bool,isbase64:bool):
        if sms == True: 
             CSVstring = convert_to_ascii(CSVstring)
             lines = CSVstring.split('\n')
        else: lines = [x[0] for x in re.findall(r'((?:(\"+)[\s\S]+?\2|[^"\n]+)+)', CSVstring)]
        csv_binary = io.StringIO()
        csv_writer = csv.writer(csv_binary, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        # Iterate through each line and write it as a CSV row
        for line in lines:
        # Remove any leading/trailing whitespace
            line = line.strip()

        # Skip empty lines
            if not line:
                continue

        # Use StringIO to create a file-like object for CSV parsing
            line_io = io.StringIO(line)
            csv_reader = csv.reader(line_io, delimiter=',', quotechar='"')

        # Extract and write the fields as a CSV row
            for row in csv_reader:
                csv_writer.writerow(row)
        csv_binary.seek(0)  
        Converted_xlsx_binary = BytesIO()
        read_file_product = pd.read_csv(csv_binary,encoding='utf8',dtype=str)
        read_file_product.to_excel(Converted_xlsx_binary, index = None, header=True)
        Converted_xlsx_binary.seek(0)  
        if isbase64 == True: return base64.b64encode(Converted_xlsx_binary.read()).decode('ascii')
        else: return Converted_xlsx_binary

def vld1(inputtime): #time validate
    a = str(inputtime)
    if a == None or a=="" or a =="None" or a == "NaN" or a == "nan":
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

def vld2(mode,date_string):
    if mode == 1:
        pattern = r"^(0?[1-9]|1[012])/((19|20)\d\d)$"
        if re.match(pattern, date_string):
            return str("01/"+date_string)
        else:
            return "01/01/1900"
    elif mode == 2:
        pattern = r"^(19|20)\d{2}$"
        if re.match(pattern, date_string):
            return str("01/01/"+date_string)
        else:
            return "01/01/1900"
    else:
        pattern = r"^(0?[1-9]|[12][0-9]|3[01])/(0?[1-9]|1[012])/((19|20)\d\d)$"
        if re.match(pattern, date_string):
            return str(date_string)
        else:
            return "01/01/1900" 

app = FastAPI()
async def verify_token(request: Request):
    mb_token = request.headers.get('mb-token')
    if mb_token is None or mb_token != "QS0R8d3b9q_GXIgdyDqCE4u5YV3yw1UliQYq22m4x5E=":
        raise HTTPException(status_code=401, detail="Unauthorized")

@app.get('/')
async def reject(token_verified: bool = Depends(verify_token)):
    return

@app.get('/pingtestfromlocalsupervisor')
async def ping():
    try:
        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except: raise HTTPException(status_code=500, detail="error")

#--------------------------------------------
@app.post('/mbcv')
async def convertcv(inputdata: dict, token_verified: bool = Depends(verify_token)):
    timestamp = int(datetime.now().timestamp())
    kill()
    try:
        print("1 - ",datetime.now()) 
        bytestring = inputdata.get("bytestring","")
        binary_data = base64.b64decode(bytestring)
        print("11 - ",datetime.now()) 
        with open(f"C:\\MB\\CV\\{timestamp}.xlsx", "wb") as file:
            file.write(binary_data)   
        print("12 - ",datetime.now()) 
        excel = win32.gencache.EnsureDispatch('Excel.Application',pythoncom.CoInitialize())
        excel.Interactive = True
        cv_type = 'new'
        fanpage = social = MBemp = newspaper = newssite = website = relative = linkedin = vnwork = topcv = topdev = itviec = ybox = careerbuilder =  "0"
        print("open wb - ",datetime.now()) 
        wb = excel.Workbooks.Open(f"C:\\MB\\CV\\{timestamp}.xlsx")
        print("done open wb - ",datetime.now()) 
        if wb.ActiveSheet.Name != "HRRF02":
            return JSONResponse(content={}, status_code=202) #detail="Bạn đang không dùng CV theo mẫu của MB")
        ws = wb.Worksheets('HRRF02')
        #Convert CV => Image
        print("start CV - ",datetime.now()) 
        CvImg = BytesIO()
        ws.Range(ws.Cells(1,1),ws.Cells(111,54)).CopyPicture(Format = 2)
        img_cv = ImageGrab.grabclipboard()
        img_cv.save(CvImg,'JPEG')
        CvImg.seek(0) 
        CvImgBase64 = base64.b64encode(CvImg.read()).decode('ascii')
        CvImg.seek(0)   
        print("end cv - ",datetime.now())    
        #Get profile pic
        PpImg = BytesIO()
        ws.Range(ws.Cells(2,37),ws.Cells(11,43)).CopyPicture(Format = 2)
        img_pp = ImageGrab.grabclipboard()
        img_pp.save(PpImg,'JPEG')
        PpImg.seek(0) 
        ProfilePicBase64 = base64.b64encode(PpImg.read()).decode('ascii')
        PpImg.seek(0)    
        if cv_type == 'new':
            print('Converting data to JSON')
            if ws.Range('N83').Value == "Website tuyển dụng của MB": website = "1"
            elif ws.Range('N83').Value == "Fanpage Tuyển dụng MBBank": fanpage = "2"
            elif ws.Range('N83').Value == "Cán bộ nhân viên MB giới thiệu": MBemp = "5"
            elif ws.Range('N83').Value == "Bạn bè/Người thân": relative = "4"
            elif ws.Range('N83').Value == "LinkedIn": linkedin = "8"
            elif ws.Range('N83').Value == "Vietnamwork": vnwork = "9"
            elif ws.Range('N83').Value == "TopCV": topcv = "10"
            elif ws.Range('N83').Value == "TopDev": topdev = "11"
            elif ws.Range('N83').Value == "ITViec":itviec = "12"
            elif ws.Range('N83').Value == "YBox": ybox = "13"
            elif ws.Range('N83').Value == "CareerBuilder": careerbuilder = "14"
            print(ws.Range('Q18').Value)
            print(null_vld(ws.Range('Q18').Value))
            print(null_vld(ws.Range('Q18').Value) +"/"+ null_vld(ws.Range('S18').Value)+"/"+null_vld(ws.Range('U18').Value))
            print(vld(0,null_vld(ws.Range('Q18').Value) +"/"+ null_vld(ws.Range('S18').Value)+"/"+null_vld(ws.Range('U18').Value)))
            data = {"CV_image": CvImgBase64,
                                "profile_picture": ProfilePicBase64,
                                "NV1" :null_vld(ws.Range('B13').Value),
                                "NV2" : null_vld(ws.Range('P13').Value),
                                "DiaDiemLamViec" : null_vld(ws.Range('AJ13').Value),

                                "1_HovaTen" : null_vld(ws.Range('B18').Value.title().strip()),
                                "1_NgaySinh" :  vld(0,null_vld(ws.Range('Q18').Value) +"/"+ null_vld(ws.Range('S18').Value)+"/"+null_vld(ws.Range('U18').Value)),
                                "1_GioiTinh" : null_vld(ws.Range('Y18').Value),
                                "1_ChieuCao" : null_vld(ws.Range('AG18').Value),
                                "1_CanNang" : null_vld(ws.Range('AM18').Value),

                                "1_thuongtru_tinh" : null_vld(ws.Range('B20').Value),
                                "1_thuongtru_quan" : null_vld(ws.Range('J20').Value),
                                "1_thuongtru_diachi" : null_vld(ws.Range('R20').Value),
                                "1_cell_personal" : null_vld(ws.Range('AG20').Value),

                                "1_NoiO_tinh" : null_vld(ws.Range('B22').Value),
                                "1_NoiO_quan" : null_vld(ws.Range('J22').Value),
                                "1_NoiO_diachi" : null_vld(ws.Range('R22').Value),
                                "1_TtHonNhan" : null_vld(ws.Range('AG22').Value),

                                "1_Email" : null_vld(ws.Range('B24').Value),
                                "1_Cmnd_passport" : null_vld(ws.Range('Q24').Value),
                                "1_NoiCap" : null_vld(ws.Range('Y24').Value),
                                "1_NgayCap" :  vld(0,null_vld(ws.Range('AG24').Value) +"/"+ null_vld(ws.Range('AH24').Value)+"/"+null_vld(ws.Range('AI24').Value)),
                                "1_DanToc" : null_vld(ws.Range('AM24').Value),

                                "1_UrgentTen" : null_vld(ws.Range('B26').Value),
                                "1_UrgentSdt" :  null_vld(ws.Range('AG26').Value),
                                "1_UrgentQuanHe" : null_vld(ws.Range('AM26').Value),
                                "1_NgayBatDau" :  null_vld(ws.Range('B28').Value),
                                "1_ThoiDiemLienHe" : null_vld(ws.Range('M28').Value),
                                "1_MucLuongDeNghi" : null_vld(ws.Range('AG28').Value),

                                "2_Ten_THPT" : null_vld(ws.Range('B33').Value),
                                "2_ChuyenNganh_THPT" : null_vld(ws.Range('Q33').Value),
                                "2_Thoigian_start_THPT" : vld(2,null_vld(ws.Range('Y33').Value)),
                                "2_Thoigian_end_THPT" : vld(1,null_vld(ws.Range('AC33').Value)+ "/"+null_vld(ws.Range('AD33').Value)),
                                "2_DiemTrungBinh_THPT" : null_vld(ws.Range('AG33').Value),
                                "2_LoaiHinhDaoTao_THPT" : null_vld(ws.Range('AP33').Value),

                                "2_Ten_TC" : null_vld(ws.Range('B34').Value),
                                "2_ChuyenNganh_TC" : null_vld(ws.Range('Q34').Value),
                                "2_Thoigian_start_TC" : vld(2,null_vld(ws.Range('Y34').Value)),
                                "2_Thoigian_end_TC" : vld(1,null_vld(ws.Range('AC34').Value)+ "/"+null_vld(ws.Range('AD34').Value)),
                                "2_DiemTrungBinh_TC" : null_vld(ws.Range('AG34').Value),
                                "2_LoaiHinhDaoTao_TC" : null_vld(ws.Range('AP34').Value),

                                "2_Ten_CD" : null_vld(ws.Range('B35').Value),
                                "2_ChuyenNganh_CD" : null_vld(ws.Range('Q35').Value),
                                "2_Thoigian_start_CD" : vld(2,null_vld(ws.Range('Y35').Value)),
                                "2_Thoigian_end_CD" : vld(1,null_vld(ws.Range('AC35').Value)+ "/"+null_vld(ws.Range('AD35').Value)),
                                "2_DiemTrungBinh_CD" : null_vld(ws.Range('AG35').Value),
                                "2_LoaiHinhDaoTao_CD" : null_vld(ws.Range('AP35').Value),

                                "2_Ten_DH" : null_vld(ws.Range('B36').Value),
                                "2_ChuyenNganh_DH" : null_vld(ws.Range('Q36').Value),
                                "2_Thoigian_start_DH" : vld(2,null_vld(ws.Range('Y36').Value)),
                                "2_Thoigian_end_DH" : vld(1,null_vld(ws.Range('AC36').Value)+ "/"+null_vld(ws.Range('AD36').Value)),
                                "2_DiemTrungBinh_DH" : null_vld(ws.Range('AG36').Value),
                                "2_LoaiHinhDaoTao_DH" : null_vld(ws.Range('AP36').Value),

                                "2_Ten_TDH" : null_vld(ws.Range('B37').Value),
                                "2_ChuyenNganh_TDH" : null_vld(ws.Range('Q37').Value),
                                "2_Thoigian_start_TDH" : vld(2,null_vld(ws.Range('Y37').Value)),
                                "2_Thoigian_end_TDH" : vld(1,null_vld(ws.Range('AC37').Value)+ "/"+null_vld(ws.Range('AD37').Value)),
                                "2_DiemTrungBinh_TDH" : null_vld(ws.Range('AG37').Value),
                                "2_LoaiHinhDaoTao_TDH" : null_vld(ws.Range('AP37').Value),

                                "2_TiengAnh_Nghe" : null_vld(ws.Range('U39').Value),
                                "2_TiengAnh_Noi" : null_vld(ws.Range('AA39').Value),
                                "2_TiengAnh_Doc" : null_vld(ws.Range('AG39').Value),
                                "2_TiengAnh_Viet" : null_vld(ws.Range('AM39').Value),

                                "2_NNK_Ten" : null_vld(ws.Range('G41').Value),
                                "2_NNK_Nghe" : null_vld(ws.Range('U41').Value),
                                "2_NNK_Noi" : null_vld(ws.Range('AA41').Value),
                                "2_NNK_Doc" : null_vld(ws.Range('AG41').Value),
                                "2_NNK_Viet" : null_vld(ws.Range('AM41').Value),

                                "2_ChungChi_Toeic" : null_vld(ws.Range('H43').Value),
                                "2_ChungChi_Toefl" : null_vld(ws.Range('N43').Value),
                                "2_ChungChi_IELTS" : null_vld(ws.Range('T43').Value),
                                "2_ChungChi_Khac" : null_vld(ws.Range('Z43').Value),
                                "2_ChungChi_NgayThi" :  vld(0,null_vld(ws.Range('AM43').Value) +"/"+ null_vld(ws.Range('AN43').Value)+"/"+null_vld(ws.Range('AO43').Value)),


                                "2_Giaithuong" : null_vld(ws.Range('B45').Value),
                                "2_ThongTinThemVeGiaiThuong" : null_vld(ws.Range('I45').Value),
                                "2_ThongTinKhac" : null_vld(ws.Range('T45').Value),
                                "2_VaiTroToChucXaHoi" : null_vld(ws.Range('AI45').Value),


                                "3_DonViCongTac_1" : null_vld(ws.Range('B51').Value),
                                "3_ChucVu_1" : null_vld(ws.Range('L51').Value),
                                "3_ThoiGian_Start_1" :  vld(1,null_vld(ws.Range('Q51').Value) +"/" +null_vld(ws.Range('R51').Value)),
                                "3_ThoiGian_End_1" :  vld(1,null_vld(ws.Range('U51').Value) +"/" +null_vld(ws.Range('V51').Value)),
                                "3_QuanLy_1" : null_vld(ws.Range('Y51').Value),
                                "3_CongViecPhuTrach_1" : null_vld(ws.Range('AD51').Value),
                                "3_LyDoThoiViec_1" : null_vld(ws.Range('AM51').Value),

                                "3_DonViCongTac_2" : null_vld(ws.Range('B52').Value),
                                "3_ChucVu_2" : null_vld(ws.Range('L52').Value),
                                "3_ThoiGian_Start_2" :  vld(1,null_vld(ws.Range('Q52').Value) +"/" +null_vld(ws.Range('R52').Value)),
                                "3_ThoiGian_End_2" :  vld(1,null_vld(ws.Range('U52').Value) +"/" +null_vld(ws.Range('V52').Value)),
                                "3_QuanLy_2" : null_vld(ws.Range('Y52').Value),
                                "3_CongViecPhuTrach_2" : null_vld(ws.Range('AD52').Value),
                                "3_LyDoThoiViec_2" : null_vld(ws.Range('AM52').Value),

                                "3_DonViCongTac_3" : null_vld(ws.Range('B53').Value),
                                "3_ChucVu_3" : null_vld(ws.Range('L53').Value),
                                "3_ThoiGian_Start_3" :  vld(1,null_vld(ws.Range('Q53').Value)  +"/" +null_vld(ws.Range('R53').Value)),
                                "3_ThoiGian_End_3" :  vld(1,null_vld(ws.Range('U53').Value)  +"/" +null_vld(ws.Range('V53').Value)),
                                "3_QuanLy_3" : null_vld(ws.Range('Y53').Value),
                                "3_CongViecPhuTrach_3" : null_vld(ws.Range('AD53').Value),
                                "3_LyDoThoiViec_3" : null_vld(ws.Range('AM53').Value),

                                "3_ThoiGian_Start_4" :  vld(0,null),
                                "3_ThoiGian_End_4" :  vld(0,null),


                                "3_DonViCongTacConLai_1" : null_vld(ws.Range('B57').Value),
                                "3_ChucVuConLai_1" : null_vld(ws.Range('L57').Value),
                                "3_ThoiGianConLai_Start_1" :  vld(1,null_vld(ws.Range('Q57').Value)  +"/"+null_vld(ws.Range('R57').Value)),
                                "3_ThoiGianConLai_End_1" :      vld(1,null_vld(ws.Range('U57').Value)  +"/"+null_vld(ws.Range('V57').Value)),                           
                                "3_CongViecPhuTrachConLai_1" : null_vld(ws.Range('Y57').Value),

                                "3_DonViCongTacConLai_2" : null_vld(ws.Range('B58').Value),
                                "3_ChucVuConLai_2" : null_vld(ws.Range('L58').Value),
                                "3_ThoiGianConLai_Start_2" :  vld(1,null_vld(ws.Range('Q58').Value)  +"/"+ null_vld(ws.Range('R58').Value)),
                                "3_ThoiGianConLai_End_2" :  vld(1,null_vld(ws.Range('U58').Value)  +"/"+null_vld(ws.Range('V58').Value)),
                                "3_CongViecPhuTrachConLai_2" : null_vld(ws.Range('Y58').Value),

                                "3_DonViCongTacConLai_3" : null_vld(ws.Range('B59').Value),
                                "3_ChucVuConLai_3" : null_vld(ws.Range('L59').Value),
                                "3_ThoiGianConLai_Start_3" :  vld(1,null_vld(ws.Range('Q59').Value)  +"/"+null_vld(ws.Range('R59').Value)),
                                "3_ThoiGianConLai_End_3" : vld(1,null_vld(ws.Range('U59').Value)  +"/"+null_vld(ws.Range('V59').Value)),                                
                                "3_CongViecPhuTrachConLai_3" : null_vld(ws.Range('Y59').Value),

                                "3_ThanhTichNoiBat" : null_vld(ws.Range('B61').Value),
                                "3_MucLuongGanNhat" : null_vld(ws.Range('AK61').Value),


                                "4_HovaTen_1" : null_vld(ws.Range('B65').Value),
                                "4_QuanHe_1" : null_vld(ws.Range('L65').Value),
                                "4_NamSinh_1" :null_vld(ws.Range('Q65').Value), 
                                "4_GioiTinh_1" : null_vld(ws.Range('V65').Value),
                                "4_NgheNghiep_NoiLV_1" : null_vld(ws.Range('AA65').Value),

                                "4_HovaTen_2" : null_vld(ws.Range('B66').Value),
                                "4_QuanHe_2" : null_vld(ws.Range('L66').Value),
                                "4_NamSinh_2" : null_vld(ws.Range('Q66').Value),
                                "4_GioiTinh_2" : null_vld(ws.Range('V66').Value),
                                "4_NgheNghiep_NoiLV_2" : null_vld(ws.Range('AA66').Value),

                                "4_HovaTen_3" : null_vld(ws.Range('B67').Value),
                                "4_QuanHe_3" : null_vld(ws.Range('L67').Value),
                                "4_NamSinh_3" : null_vld(ws.Range('Q67').Value),
                                "4_GioiTinh_3" : null_vld(ws.Range('V67').Value),
                                "4_NgheNghiep_NoiLV_3" : null_vld(ws.Range('AA67').Value),

                                "4_HovaTen_4" : null_vld(ws.Range('B68').Value),
                                "4_QuanHe_4" : null_vld(ws.Range('L68').Value),
                                "4_NamSinh_4" : null_vld(ws.Range('Q68').Value),
                                "4_GioiTinh_4" : null_vld(ws.Range('V68').Value),
                                "4_NgheNghiep_NoiLV_4" : null_vld(ws.Range('AA68').Value),
                                "4_HovaTen_5" : null_vld(ws.Range('B69').Value),
                                "4_QuanHe_5" : null_vld(ws.Range('L69').Value),
                                "4_NamSinh_5" : null_vld(ws.Range('Q69').Value),
                                "4_GioiTinh_5" : null_vld(ws.Range('V69').Value),
                                "4_NgheNghiep_NoiLV_5" : null_vld(ws.Range('AA69').Value),
                                "4_HovaTen_6" : null_vld(ws.Range('B70').Value),
                                "4_QuanHe_6" : null_vld(ws.Range('L70').Value),
                                "4_NamSinh_6" : null_vld(ws.Range('Q70').Value),
                                "4_GioiTinh_6" :  null_vld(ws.Range('V70').Value),
                                "4_NgheNghiep_NoiLV_6" : null_vld(ws.Range('AA70').Value),

                                "5_PhamChatNangKhieu" : null_vld(ws.Range('B74').Value),
                                "5_KyNangDacBiet" : null_vld(ws.Range('L74').Value),
                                "5_ThanhVien" : null_vld(ws.Range('V74').Value),
                                "5_KhoaHoc" : null_vld(ws.Range('AG74').Value),

                                "6_DanhHieu" : null_vld(ws.Range('B79').Value),
                                "6_HinhThucKyLuat" : null_vld(ws.Range('X79').Value),
                                "6_ThoiGianKyLuat" :  vld(0,null_vld(ws.Range('AM79').Value) +"/"+ null_vld(ws.Range('AN79').Value)+"/"+null_vld(ws.Range('AO79').Value)),

                                "7_website" :  website,
                                "7_fanpage" :  fanpage,
                                "7_BanBeNguoiThan" : relative,
                                "7_MB" :  MBemp,
                                "7_linkedin": linkedin,
                                "7_vietnamwork": vnwork,
                                "7_topcv": topcv,
                                "7_topdev": topdev,
                                "7_itviec": itviec,
                                "7_ybox": ybox,
                                "7_careerbuilder": careerbuilder,
                                "7_Khac" : null_vld(ws.Range('AE83').Value),

                                "8_HovaTen_1" : null_vld(ws.Range('B87').Value),
                                "8_ChucVu_1" : null_vld(ws.Range('L87').Value),
                                "8_DonViCongTac_1" : null_vld(ws.Range('Q87').Value),
                                "8_MoiQuanHe_1" : null_vld(ws.Range('Y87').Value),
                                "8_Mobile_1" : null_vld(ws.Range('AG87').Value),
                                "8_email_1" : null_vld(ws.Range('AM87').Value), 

                                "8_HovaTen_2" : null_vld(ws.Range('B88').Value),
                                "8_ChucVu_2" : null_vld(ws.Range('L88').Value),
                                "8_DonViCongTac_2" : null_vld(ws.Range('Q88').Value),
                                "8_MoiQuanHe_2" : null_vld(ws.Range('Y88').Value),
                                "8_Mobile_2" : null_vld(ws.Range('AG88').Value),
                                "8_email_2" : null_vld(ws.Range('AM88').Value),

                                "8_HovaTen_3" : null_vld(ws.Range('B89').Value),
                                "8_ChucVu_3" : null_vld(ws.Range('L89').Value),
                                "8_DonViCongTac_3" : null_vld(ws.Range('Q89').Value),
                                "8_MoiQuanHe_3" : null_vld(ws.Range('Y89').Value),
                                "8_Mobile_3" : null_vld(ws.Range('AG89').Value),
                                "8_email_3" : null_vld(ws.Range('AM89').Value),

                                "9_HovaTen_1" : null_vld(ws.Range('B93').Value),
                                "9_ChucVu_1" : null_vld(ws.Range('L93').Value),
                                "9_DonViCongTac_1" : null_vld(ws.Range('Q93').Value),
                                "9_MoiQuanHe_1" : null_vld(ws.Range('Y93').Value),
                                "9_Mobile_1" : null_vld(ws.Range('AG93').Value),
                                "9_email_1" : null_vld(ws.Range('AM93').Value),

                                "9_HovaTen_2" : null_vld(ws.Range('B94').Value),
                                "9_ChucVu_2" : null_vld(ws.Range('L94').Value),
                                "9_DonViCongTac_2" : null_vld(ws.Range('Q94').Value),
                                "9_MoiQuanHe_2" : null_vld(ws.Range('Y94').Value),
                                "9_Mobile_2" : null_vld(ws.Range('AG94').Value),
                                "9_email_2" : null_vld(ws.Range('AM94').Value),

                                "9_HovaTen_3" : null_vld(ws.Range('B95').Value),
                                "9_ChucVu_3" : null_vld(ws.Range('L95').Value),
                                "9_DonViCongTac_3" : null_vld(ws.Range('Q95').Value),
                                "9_MoiQuanHe_3" : null_vld(ws.Range('Y95').Value),
                                "9_Mobile_3" : null_vld(ws.Range('AG95').Value),
                                "9_email_3" : null_vld(ws.Range('AM95').Value),

                                "10_ViTriTungUngTuyen" : null_vld(ws.Range('B99').Value),
                                "10_ThoiGian" :  vld(0,null_vld(ws.Range('L99').Value) +"/"+ null_vld(ws.Range('N99').Value)+"/"+null_vld(ws.Range('P99').Value)),
                                "10_KhiemKhuyetNgoaiHinh_Co" : null_vld(ws.Range('AG99').Value),
                                "10_KhiemKhuyetNgoaiHinh_Khong" : null_vld(ws.Range('AL99').Value),

                                "10_DiemManh" : null_vld(ws.Range('B101').Value),
                                "10_DiemYeu" : null_vld(ws.Range('V101').Value),
                                "10_LyDo_1" : null_vld(ws.Range('D103').Value),
                                "10_LyDo_2" : null_vld(ws.Range('D104').Value),
                                "10_LyDo_3" : null_vld(ws.Range('D105').Value),
                                "10_DienGiaiLyDo" : null_vld(ws.Range('Q103').Value),
                                "10_KhaNangDacBiet": null_vld(ws.Range('Q107').Value),
                                "10_NgayTao" :  vld(0,null_vld(ws.Range('B110').Value) +"/"+ null_vld(ws.Range('D110').Value)+"/"+null_vld(ws.Range('F110').Value))      
                                }
            print("4 - ",datetime.now()) 
            kill()

        return JSONResponse(content=data, status_code=200)
    except Exception as error:
        print(error)
        raise HTTPException(status_code=500, detail="error")

#---------
@app.post('/dxpvnb')
async def dexuatphongvan(inputdata: dict, token_verified: bool = Depends(verify_token)):
    try:
        body = inputdata.get("body","")
        binary_data = base64.b64decode(body)
        dxpv_binary = BytesIO()
        dxpv_binary.write(binary_data)
        df = pd.read_excel(dxpv_binary)
        data = []
        for index, row in df.iterrows():
            data.append({"CV_layout": "importdxpv",
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
                "NgaySinh": vld1(row.iloc[11]) })
        return JSONResponse(content=data, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

#--------------------------------------------
@app.post('/sms')
async def sms(smsdata: dict, token_verified: bool = Depends(verify_token)):
    try:
        CSVstring = smsdata.get("body","")
        xlsxfile = csv2xlsx(CSVstring,True,True)
        datasms = {"file":xlsxfile}
        return JSONResponse(content=datasms, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

#--------------------------------------------
@app.post('/hcm')
async def exportHCM(smsdata: dict, token_verified: bool = Depends(verify_token)):
    try:
        CSVstring = smsdata.get("body","")
        filename = smsdata.get("filename","")
        lines = [x[0] for x in re.findall(r'((?:(\"+)[\s\S]+?\2|[^"\n]+)+)', CSVstring)]
        with open(r"C:\\MB\\HCM\\"+filename, mode='w', newline='',encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            # Iterate through each line and write it as a CSV row
            for line in lines:
            # Remove any leading/trailing whitespace
                line = line.strip()

            # Skip empty lines
                if not line:
                    continue

            # Use StringIO to create a file-like object for CSV parsing
                line_io = io.StringIO(line)
                csv_reader = csv.reader(line_io, delimiter=',', quotechar='"')

            # Extract and write the fields as a CSV row
                for row in csv_reader:
                    csv_writer.writerow(row)
        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

#--------------------------------------------
@app.post('/thkqtd')
async def thkqtd(dsuv: dict, token_verified: bool = Depends(verify_token)):
    try:
        CSVstring = dsuv.get("body","")
        filename = dsuv.get("filename","")
        lines = [x[0] for x in re.findall(r'((?:(\"+)[\s\S]+?\2|[^"\n]+)+)', CSVstring)]
        with open(r"C:\\MB\\THKQTD\\"+filename, mode='w', newline='',encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            # Iterate through each line and write it as a CSV row
            for line in lines:
            # Remove any leading/trailing whitespace
                line = line.strip()

            # Skip empty lines
                if not line:
                    continue

            # Use StringIO to create a file-like object for CSV parsing
                line_io = io.StringIO(line)
                csv_reader = csv.reader(line_io, delimiter=',', quotechar='"')

            # Extract and write the fields as a CSV row
                for row in csv_reader:
                    csv_writer.writerow(row)
        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/dxluong')
async def dxl(dxl: dict, token_verified: bool = Depends(verify_token)):
    try:
        CSVstring = dxl.get("body","")
        data = csv2xlsx(CSVstring,False,True)
        return JSONResponse(content={"file": data}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

#--------------------------------------------
@app.post('/dstv')
async def exportDS_ThiViet(dstv: dict, token_verified: bool = Depends(verify_token)):
    try:
        CSVstring = dstv.get("body","")
        data = csv2xlsx(CSVstring,False,True)
        return JSONResponse(content={"file":data}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/dsuv') #Export Theo mau Database.xlsm
async def exportDSUV(dsuv: dict, token_verified: bool = Depends(verify_token)):
    try:
        CSVstring = dsuv.get("body","")
        data = csv2xlsx(CSVstring,False,True)
        return JSONResponse(content={"file":data}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/docx') #scan follina
async def docx(docxfile: dict, token_verified: bool = Depends(verify_token)):
    try:
        bytestring = docxfile.get("body","")
        binary_data = base64.b64decode(bytestring)
        # Open the .docx file as a zip file
        target_string = 'TargetMode'
        with zipfile.ZipFile(BytesIO(binary_data), 'r') as zip_file:
            # Extract the contents of the document.xml.rels file
            rels_content = zip_file.read('word/_rels/document.xml.rels').decode('utf-8')
            if target_string in rels_content:
                return JSONResponse(content={"message": "Eh yooo, get out!"}, status_code=417)
            else: return JSONResponse(content={"message": "OK"}, status_code=200)     
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/pdfscan') #scan pdf
async def pdf(pdf_file: dict, token_verified: bool = Depends(verify_token)):
    try:
        bytestring = pdf_file.get("body","")
        binary_data = base64.b64decode(bytestring)
        # Open the .docx file as a zip file
        vulnerable_strings = [b'/OpenAction', b'/AA', b'/JavaScript', b'/JS', b'/AcroForm', b'/XFA', b'/URI', b'/SubmitForm', b'/GoToR', b'CBSharedReviewIfOfflineDialog']
        if any(s in binary_data for s in vulnerable_strings):
        # Potential vulnerability detected
            return JSONResponse(content={"message": "Eh yooo, get out!"}, status_code=417)
        else: return JSONResponse(content={"message": "OK"}, status_code=200)     
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/xlsx') #scan follina
async def xlsx(xlsxfile: dict, token_verified: bool = Depends(verify_token)):
    try:
        bytestring = xlsxfile.get("body","")
        binary_data = base64.b64decode(bytestring)
        # Open the .docx file as a zip file
        target_string = 'HRRF02'
        with zipfile.ZipFile(BytesIO(binary_data), 'r') as zip_file:
            # Extract the contents of the document.xml.rels file
            rels_content = zip_file.read('xl/workbook.xml').decode('utf-8')
            if target_string not in rels_content:
                return JSONResponse(content={"message": "Eh yooo, get out!"}, status_code=417)
            else: return JSONResponse(content={"message": "OK"}, status_code=200)     
    except:
        raise HTTPException(status_code=500, detail="error")

#--------------------------------------------
@app.post('/cloneimage') 
async def image_clone(data: dict, token_verified: bool = Depends(verify_token)):
    try:
        body = data.get("body","")
        filetype = data.get("filetype","")
        # Decode the base64 string into binary data
        binary_data = base64.b64decode(body)

        # Open the image from binary data
        original_image = Image.open(BytesIO(binary_data))

        # Clone or copy the image
        cloned_image = original_image.copy()

        # Save the cloned image to BytesIO
        cloned_image_data = BytesIO()
        cloned_image.save(cloned_image_data, format=filetype)  # You can change the format as needed

        # Convert the cloned image data to base64
        cloned_base64 = base64.b64encode(cloned_image_data.getvalue()).decode("ascii")

        return JSONResponse(content={"message": cloned_base64}, status_code=200)
    except Exception as e:
        print(e)
        raise HTTPException(status_code=500, detail=f"{e}")


#--------------------------------------------
@app.post('/dxtd') 
async def dxtd(data: dict, token_verified: bool = Depends(verify_token)):
    try:
        body = data.get("body","")
        binary_data = base64.b64decode(body)
        dxtd_binary = BytesIO()
        dxtd_binary.write(binary_data)
        df = pd.read_excel(dxtd_binary)
        data = []
        for index, row in df.iterrows():
            data.append({
            "CV_layout": "importdxtd",
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
            })
            print(f'Collecting proposal data #{index+1} to Dataverse')
        return JSONResponse(content=data, status_code=200)
    except Exception as e:
        print(e)
        raise HTTPException(status_code=500, detail=f"{e}")

#--------------------------------------------
@app.post('/importdsuv') 
async def importDSUV(data: dict, token_verified: bool = Depends(verify_token)):
    try:
        body = data.get("body","")
        binary_data = base64.b64decode(body)
        dxtd_binary = BytesIO()
        dxtd_binary.write(binary_data)
        df = pd.read_excel(dxtd_binary)
        data = []
        for index, row in df.iterrows():
            Ten_THPT = ChuyenNganh_THPT = Thoigian_end_THPT =   Ten_TC = ChuyenNganh_TC = Thoigian_end_TC = Ten_CD = ChuyenNganh_CD = Thoigian_end_CD = Ten_DH = ChuyenNganh_DH = Thoigian_end_DH = Ten_TDH = ChuyenNganh_TDH = Thoigian_end_TDH = str(None)
            if str(row.iloc[5]) == "Phổ thông trung học":
                    Ten_THPT = str(row.iloc[6])
                    ChuyenNganh_THPT = str(row.iloc[7])
                    Thoigian_end_THPT = vld2(2,str(row.iloc[8]))

            elif str(row.iloc[5]) == "Trung cấp":
                    Ten_TC = str(row.iloc[6])
                    ChuyenNganh_TC = str(row.iloc[7])
                    Thoigian_end_TC = vld2(2,str(row.iloc[8]))

            elif str(row.iloc[5]) == "Cao đẳng":
                    Ten_CD = str(row.iloc[6])
                    ChuyenNganh_CD = str(row.iloc[7])
                    Thoigian_end_CD = vld2(2,str(row.iloc[8]))

            elif str(row.iloc[5]) == "Đại học":
                    Ten_DH = str(row.iloc[6])
                    ChuyenNganh_DH = str(row.iloc[7])
                    Thoigian_end_DH = vld2(2,str(row.iloc[8]))

            elif str(row.iloc[5]) == "Trên ĐH":
                    Ten_TDH = str(row.iloc[6])
                    ChuyenNganh_TDH = str(row.iloc[7])
                    Thoigian_end_TDH = vld2(2,str(row.iloc[8]))
            print(str(row.iloc[3]))
            print(Thoigian_end_DH)

            data.append({ 
        "CV_layout": "db",
        "1_HovaTen" : str(row.iloc[0]),
        "1_NgaySinh" :  vld2(0,str(row.iloc[3])),
        "1_GioiTinh" : str(row.iloc[4]),
        "1_ChieuCao" : str(row.iloc[15]),
        "1_CanNang" : str(row.iloc[16]),


        "1_thuongtru_tinh" : str(row.iloc[21]),
        "1_thuongtru_quan" : str(row.iloc[20]),
        "1_thuongtru_diachi" : str(row.iloc[19]),
        "1_cell_personal" : str(int(str(row.iloc[2]).replace(" ",""))),

        "1_NoiO_tinh" : str(row.iloc[24]),
        "1_NoiO_quan" : str(row.iloc[23]),
        "1_NoiO_diachi" : str(row.iloc[22]),
        "1_TtHonNhan" : str(row.iloc[17]),

        "1_Email" : str(row.iloc[1]),
        "1_Cmnd_passport" : str(row.iloc[12]),
        "1_NoiCap" : str(row.iloc[13]),
        "1_NgayCap" :  vld2(0,str(row.iloc[14])),
        "1_DanToc" : str(row.iloc[18]),


        "2_Ten_THPT" : Ten_THPT,
        "2_ChuyenNganh_THPT" : ChuyenNganh_THPT,
        "2_Thoigian_end_THPT" : vld2(0,Thoigian_end_THPT),


        "2_Ten_TC" : Ten_TC,
        "2_ChuyenNganh_TC" : ChuyenNganh_TC,
        "2_Thoigian_end_TC" : vld2(0,Thoigian_end_TC),


        "2_Ten_CD" : Ten_CD,
        "2_ChuyenNganh_CD" : ChuyenNganh_CD,
        "2_Thoigian_end_CD" : vld2(0,Thoigian_end_CD),

        "2_Ten_DH" : Ten_DH,
        "2_ChuyenNganh_DH" : ChuyenNganh_DH,
        "2_Thoigian_end_DH" : vld2(0,Thoigian_end_DH),

        "2_Ten_TDH" : Ten_TDH,
        "2_ChuyenNganh_TDH" : ChuyenNganh_TDH,
        "2_Thoigian_end_TDH" : vld2(0,Thoigian_end_TDH),


        "10_DiemManh" : str(row.iloc[9]),
        "10_DiemYeu" : str(row.iloc[10]),

        "10_DienGiaiLyDo" : str(row.iloc[11]),
        "1_NgayBatDau" :  vld2(0,null),
        "2_Thoigian_start_THPT" : vld2(0,null),
        "2_Thoigian_start_TC" : vld2(0,null),
        "2_Thoigian_start_CD" : vld2(0,null),
        "2_Thoigian_start_DH" : vld2(0,null),
        "2_Thoigian_start_TDH" : vld2(0,null),

        "2_ChungChi_NgayThi" :vld2(0,null),
        "3_ThoiGian_Start_1" :  vld2(0,null),
        "3_ThoiGian_End_1":vld2(0,null),
        "3_ThoiGian_Start_2" :  vld2(0,null),
        "3_ThoiGian_End_2":vld2(0,null),
        "3_ThoiGian_Start_3" :  vld2(0,null),
        "3_ThoiGian_End_3":vld2(0,null),
        "3_ThoiGian_Start_4" :  vld2(0,null),
        "3_ThoiGian_End_4" :  vld2(0,null),
        "3_ThoiGianConLai_Start_1" :  vld2(0,null),
        "3_ThoiGianConLai_End_1" :vld2(0,null),
        "3_ThoiGianConLai_Start_2" :  vld2(0,null),
        "3_ThoiGianConLai_End_2" :vld2(0,null),
        "3_ThoiGianConLai_Start_3" :  vld2(0,null),
        "3_ThoiGianConLai_End_3" : vld2(0,null),
        "6_ThoiGianKyLuat" :vld2(0,null),
        "10_ThoiGian" :vld2(0,null),
        "10_NgayTao":vld2(0,str(row.iloc[46]))})
            print(f'Collecting proposal data #{index+1} to Dataverse')
        return JSONResponse(content=data, status_code=200)
    except Exception as e:
        print(e)
        raise HTTPException(status_code=500, detail=f"{e}")

#--------------------------------------------
@app.post('/thhd') 
async def thhd(data: dict, token_verified: bool = Depends(verify_token)):
    try:
        #get data from dataverse
        CSVstring = data.get("THHD","")
        CSVstring_2 = data.get("THHD_Detail","")
        owner_detail = data.get("owner_detail","")
        owner_email = data.get("owner_email","")

        timestamp_int = int(datetime.now().timestamp())
        timestamp = str(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

        # Convert CSV to XLSX
        thhd_xlsx_bin = csv2xlsx(CSVstring,False,False)
        thhd_detail_xlsx_bin = csv2xlsx(CSVstring_2,False,False)
        #read template + load workbook into variable
        C = BytesIO()
        C = open("C:\\MB\\THHD_Template\\THHD.xlsx", "rb")
        C.seek(0)
        wb_c = load_workbook(C)
        C.seek(0)
        sheet_a_c = wb_c['A']
        sheet_b_c = wb_c['B']

        #create IO object
        thhd_html = io.StringIO()
        thhd_detail_html = io.StringIO()
        bytesio_a = BytesIO()
        bytesio_b = BytesIO()
        output_xlsx = BytesIO()

        # Extract data from A.xlsx and B.xlsx
        df_a = pd.read_excel(thhd_xlsx_bin)
        df_b = pd.read_excel(thhd_detail_xlsx_bin)
        # Clear existing data in sheet A and B of C.xlsx
        sheet_a_c.delete_rows(3, sheet_a_c.max_row)
        sheet_b_c.delete_rows(3, sheet_b_c.max_row)

        header_fill_color = '56637C' 

        # Write data to sheet A and B of C.xlsx
        for r_idx, row in enumerate(dataframe_to_rows(df_a, index=False), start=3):
            for c_idx, value in enumerate(row, start=1):
                sheet_a_c.cell(row=r_idx, column=c_idx, value="" if str(value) == "nan" else str(value))

        for r_idx, row in enumerate(dataframe_to_rows(df_b, index=False), start=3):
            for c_idx, value in enumerate(row, start=1):
                sheet_b_c.cell(row=r_idx, column=c_idx, value="" if str(value) == "nan" else str(value))


#format
        for sheet in wb_c:
            for row in sheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center',wrap_text=True)
                    cell.border = Border(left=Side(border_style='thin', color='00000000'),
                                        right=Side(border_style='thin', color='00000000'),
                                        top=Side(border_style='thin', color='00000000'),
                                        bottom=Side(border_style='thin', color='00000000'))
                    cell.font = Font(name= "Arial")
            for cell in sheet[3]:
                cell.fill = PatternFill(start_color=header_fill_color, end_color=header_fill_color, fill_type='solid')
                cell.font = Font(color='FFFFFFFF', bold=True)
            #save sheet A to XLSX
            if sheet.title == 'A':
                wb_c.active = wb_c['A']
                sheet['A1'] = f"Tạo bởi {owner_detail} ({owner_email}) lúc {timestamp} "
                wb_c.save(bytesio_a)

            #save sheet B to XLSX
            elif sheet.title == 'B':
                wb_c.active = wb_c['B']
                wb_c.save(bytesio_b)

        #save the whole workbook to XLSX
        
        wb_c.save(output_xlsx)
        output_xlsx.seek(0)
        bytesio_a.seek(0)
        bytesio_b.seek(0)

        #convert xlsx to HTML, each file
        xlsx2html(bytesio_a, thhd_html)
        xlsx2html(bytesio_b, thhd_detail_html)
        thhd_html.seek(0)
        thhd_detail_html.seek(0)

        #combine 2 html
        merged_content = thhd_html.getvalue() + '\n' + thhd_detail_html.getvalue()

        #write to file
        with open(f"C:\\MB\\THHD\\{timestamp_int}.html", 'w',encoding="utf-8") as output_file:
                    # Write merged content to the output file
                    output_file.write(merged_content)

        print("HTML files combined successfully!")

        options = {
            'page-size': 'Letter',
            'orientation': 'Landscape'  # Set orientation to Landscape
        }

        #convert combine HTML to PDF
        pdfkit.from_file(f"C:\\MB\\THHD\\{timestamp_int}.html", f"C:\\MB\\THHD\\{timestamp_int}.pdf",options=options)

        #convert to base64
        b64_output_xlsx = base64.b64encode(output_xlsx.read()).decode('ascii')

        with open(f"C:\\MB\\THHD\\{timestamp_int}.pdf", "rb") as xlsxbinary1:
                    b64_output_pdf = base64.b64encode(xlsxbinary1.read()).decode('ascii')
        
        data = {"thhd": b64_output_xlsx,"output_pdf":b64_output_pdf}
        return JSONResponse(content=data, status_code=200)
    except Exception as e:
        print(e)
        raise HTTPException(status_code=500, detail=f"{e}")



if __name__ == '__main__':
    print('Starting API')
    uvicorn.run(app, host='0.0.0.0', port=8080)

