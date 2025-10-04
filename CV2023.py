from openpyxl import load_workbook
import win32com.client as win32
import time
import os
import psutil
import requests
import base64
from datetime import datetime
from PIL import ImageGrab

def kill():
      for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                  proc.kill()

def null_vld(value1):
      if value1 == None:
            return ""
      else: 
            return str(value1)
# def vld(mode,date_string):
#     try:    
#         if mode == 1:
#             pattern = r"^(0?[1-9]|1[012])/((19|20)\d\d)$"
#             if re.match(pattern, date_string):
#                 return  datetime.strptime(str("01/"+date_string), "%d/%m/%Y").isoformat()
#             else:
#                 return "1900-01-01T00:00:00"
#         elif mode == 2:
#             pattern = r"^(19|20)\d{2}$"
#             if re.match(pattern, date_string):
#                 return  datetime.strptime(str("01/01/"+date_string), "%d/%m/%Y").isoformat()
#             else:
#                 return "1900-01-01T00:00:00"
#         else:
#             pattern = r"^(0?[1-9]|[12][0-9]|3[01])/(0?[1-9]|1[012])/((19|20)\d\d)$"
#             if re.match(pattern, date_string):
#                 return datetime.strptime(str(date_string), "%d/%m/%Y").isoformat()
#             else:
#                 return "1900-01-01T00:00:00" 
#     except:
#           return "1900-01-01T00:00:00" 
    
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


# def extract(p):
#         shutil.copy(f'C:\\MB\\CV\\{p}', "C:\\MB\\temp\\")
#         os.rename(f'C:\\MB\\temp\\{p}', f"C:\\MB\\temp\\{id}.zip")
#         os.mkdir(f'C:\\MB\\temp\\{id}')
#         print('created folder')
#         with zipfile.ZipFile(f'C:\\MB\\temp\\{id}.zip', 'r') as zip:
#                 zip.extractall(f'C:\\MB\\temp\\{id}')
#         print('unzipped')
#         x = os.listdir(f'C:\\MB\\temp\\{id}\\xl\\media')
#         print(len(x))
#         if len(x) >= 2:
#                 with open(f'C:\\MB\\temp\\{id}\\xl\\media\\{x[1]}', "rb") as img:
#                         pp = base64.b64encode(img.read()).decode('ascii')
#                 shutil.rmtree(f'C:\\MB\\temp\\{id}')
#                 os.rename(f'C:\\MB\\temp\\{id}.zip',r"C:\\MB\\temp\\"+p)
#                 return pp
        
#         else:
#                 shutil.rmtree(f'C:\\MB\\temp\\{id}')
#                 os.rename(f'C:\\MB\\temp\\{id}.zip',r"C:\\MB\\temp\\"+p)
#                 return None


while True:
    try:
        try:
#Preprocess--------------------------------------------------------
                
                a = os.listdir("C:\\MB\\CV")
                if len(a) > 0:
                        blank = False
                else: blank = True
                print(a)
                env_delimiter = "YYYYY"  
                delimiter = "XXXXXXXXXX"   
                envr, rest_email = a[0].split(env_delimiter, 1)
                if envr == "DEV":
                        cvupload = str("https://prod-63.southeastasia.logic.azure.com:443/workflows/5e032ca91fa0434a8f41cabcfd78bb10/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2ORhEndjRj7epCarHXReFY0c2hBrkGICYoBwfMmhjhk")
                elif envr == "UAT":
                        cvupload = str("https://prod-10.southeastasia.logic.azure.com:443/workflows/4e972afd10a546338794694201dff6a2/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jHTY5h1cTQLRWhzSJ5VMGQWAr-Bl4Z-pYRpdv2eXwzg")
                elif envr == "PROD":
                        cvupload = str("https://prod-40.southeastasia.logic.azure.com:443/workflows/70a76e6d765f4705ba487cb8faee801f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-D0IvfgsYy7o0-DxS9l_jJGQwURYNQau5_JqsWnxURA")
                owner, rest = rest_email.split(delimiter, 1)
                os.rename(r"C:\\MB\\CV\\"+a[0],r"C:\\MB\\CV\\"+rest)
                a[0] = rest
                owner = owner +"@mbbank.com.vn"
                file_extension = os.path.splitext(a[0])[1].lower()
                excel = win32.gencache.EnsureDispatch('Excel.Application')
#Convert CV ---------------------------------------------------------
                if file_extension == '.xls':
                        wbx = excel.Workbooks.Open(r"C:\\MB\\CV\\"+a[0])
                        wbx.SaveAs(f"C:\\MB\\CV\\XLS_{a[0]}x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
                        wbx.Close()
                        excel.Workbooks.Close()
                        os.remove(r"C:\\MB\\CV\\"+a[0])  
                        a[0]="XLS_"+a[0]+"x"
                        print('Converted to XLSX')
#CV Processing--------------------------------------------------------

                cv_type = ''
                fanpage = social = MBemp = newspaper = newssite = website = relative = linkedin = vnwork = topcv = topdev = itviec = ybox = careerbuilder =  "0"

                if a[0].startswith("XLS"):
                        cv_type = 'old'
                else:
                        cv_type = 'new'
                wb = excel.Workbooks.Open(r"C:\\MB\\CV\\"+a[0])
                ws = wb.Worksheets('HRRF02')
                ws.Range(ws.Cells(1,1),ws.Cells(111,54)).CopyPicture(Format = 2)
                img_cv = ImageGrab.grabclipboard()
                img_cv.save(f"C:\\MB\\{a[0]}.jpeg")
                with open(f"C:\\MB\\{a[0]}.jpeg", "rb") as cv_img:
                        CV_image = base64.b64encode(cv_img.read()).decode('ascii')
                os.remove(f"C:\\MB\\{a[0]}.jpeg")  
                #Get profile pic
                ws.Range(ws.Cells(2,37),ws.Cells(11,43)).CopyPicture(Format = 2)
                img_pp = ImageGrab.grabclipboard()
                img_pp.save(f"C:\\MB\\{a[0]}.jpeg")
                with open(f"C:\\MB\\{a[0]}.jpeg", "rb") as pp_img:
                        pp = base64.b64encode(pp_img.read()).decode('ascii')
                excel.Workbooks.Close()
                excel.Application.Quit()
                
                os.remove(f"C:\\MB\\{a[0]}.jpeg")
#                print('checkpoint 2')
        
#Openpyxl load worksheet-------------------------------------------------------
                wb1 = load_workbook(r"C:\\MB\\CV\\"+a[0])
                sheet = wb1.active
                print('Processing data of ' + sheet['B18'].value)
                id = sheet['Q24'].value
                email_8_1 = mobile_8_1 = email_8_2 = mobile_8_2 = mobile_8_3 = email_8_3 = email_9_1 = email_9_2 = email_9_3 = mobile_9_1 = mobile_9_2 = mobile_9_3 = 'None'
                print('checkpoint 2_1')
                if cv_type == 'old':
                        if sheet['B73'].value != '': website = '1'
                        if sheet['I73'].value != '': newssite = '7'
                        if sheet['P73'].value != '': newspaper = '6'
                        if sheet['W73'].value != '': relative = '4'
                        print('checkpoint 2_2')
                        if '@' in null_vld(sheet['AG77'].value): email_8_1 = sheet['AG77'].value
                        else: mobile_8_1 = sheet['AG77'].value

                        if '@' in null_vld(sheet['AG78'].value): email_8_2 = sheet['AG78'].value
                        else: mobile_8_2 = sheet['AG78'].value

                        if '@' in null_vld(sheet['AG79'].value): email_8_3 = sheet['AG79'].value
                        else: mobile_8_3 = sheet['AG79'].value

                        if '@' in null_vld(sheet['AG83'].value): email_9_1 = sheet['AG83'].value
                        else: mobile_9_1 = sheet['AG83'].value

                        if '@' in null_vld(sheet['AG84'].value): email_9_2 = sheet['AG84'].value
                        else: mobile_9_2 = sheet['AG84'].value

                        if '@' in null_vld(sheet['AG84'].value): email_9_3 = sheet['AG84'].value
                        else: mobile_9_3 = sheet['AG84'].value
                        
                        print('Converting data to JSON')
                        data = {"owner": owner,
                                "CV_image": CV_image,
                                 "username": "admin",
                                "password": "@MB2023",
                "CV_layout": "old",
                "profile_picture": pp,
                "NV1" :sheet['H13'].value,
                "NV2" : sheet['U13'].value,
                "DiaDiemLamViec" : sheet['AJ13'].value,

                "1_HovaTen" : sheet['B18'].value.title().strip(),
                "1_NgaySinh" :   vld(0,null_vld(sheet['Q18'].value) +"/"+ null_vld(sheet['S18'].value)+"/"+null_vld(sheet['U18'].value)),
                "1_GioiTinh" : sheet['Y18'].value,
                "1_ChieuCao" : sheet['AG18'].value,
                "1_CanNang" : sheet['AM18'].value,

                "1_thuongtru_diachi" : sheet['B20'].value,
                "1_cell_permanent" : sheet['AG20'].value,
                "1_cell_personal" : sheet['AM20'].value,


                "1_NoiO_diachi" : sheet['B22'].value,
                "1_TtHonNhan" : sheet['AG22'].value,

                "1_Email" : sheet['B24'].value,
                "1_Cmnd_passport" : sheet['Q24'].value,
                "1_NoiCap" : sheet['Y24'].value,
                "1_NgayCap" :  vld(0,null_vld(sheet['AG24'].value)),
                "1_DanToc" : sheet['AM24'].value,

                "1_UrgentTen" : sheet['B26'].value,
                "1_UrgentSdt" : sheet['AG26'].value,
                "1_UrgentQuanHe" : sheet['AM26'].value,

                "1_NgayBatDau" :  vld(0,null_vld(sheet['B28'].value)),
                "1_ThoiDiemLienHe" : sheet['M28'].value,
                "1_MucLuongDeNghi" : sheet['AG28'].value,

                "2_Ten_THPT" : sheet['B33'].value,
                "2_ChuyenNganh_THPT" : sheet['Q33'].value,
                "2_Thoigian_start_THPT" : vld(2,null_vld(sheet['Y33'].value)),
                "2_Thoigian_end_THPT" : vld(2,null_vld(sheet['AC33'].value)),
                "2_DiemTrungBinh_THPT" : sheet['AG33'].value,
                "2_LoaiHinhDaoTao_THPT" : sheet['AM33'].value,

                "2_Ten_TC" : sheet['B34'].value,
                "2_ChuyenNganh_TC" : sheet['Q34'].value,
                "2_Thoigian_start_TC" : vld(2,null_vld(sheet['Y34'].value)),
                "2_Thoigian_end_TC" :vld(2, sheet['A34'].value),
                "2_DiemTrungBinh_TC" : sheet['AG34'].value,
                "2_LoaiHinhDaoTao_TC" : sheet['AM34'].value,

                "2_Ten_CD" : sheet['B35'].value,
                "2_ChuyenNganh_CD" : sheet['Q35'].value,
                "2_Thoigian_start_CD" : vld(2,null_vld(sheet['Y35'].value)),
                "2_Thoigian_end_CD" :vld(2, sheet['A35'].value),
                "2_DiemTrungBinh_CD" : sheet['AG35'].value,
                "2_LoaiHinhDaoTao_CD" : sheet['AM35'].value,

                "2_Ten_DH" : sheet['B36'].value,
                "2_ChuyenNganh_DH" : sheet['Q36'].value,
                "2_Thoigian_start_DH" : vld(2,null_vld(sheet['Y36'].value)),
                "2_Thoigian_end_DH" : vld(2,null_vld(sheet['A36'].value)),
                "2_DiemTrungBinh_DH" : sheet['AG36'].value,
                "2_LoaiHinhDaoTao_DH" : sheet['AM36'].value,

                "2_Ten_TDH" : sheet['B37'].value,
                "2_ChuyenNganh_TDH" : sheet['Q37'].value,
                "2_Thoigian_start_TDH" : vld(2,null_vld(sheet['Y37'].value)),
                "2_Thoigian_end_TDH" : vld(2,null_vld(sheet['A37'].value)),
                "2_DiemTrungBinh_TDH" : sheet['AG37'].value,
                "2_LoaiHinhDaoTao_TDH" : sheet['AM37'].value,

                "2_TiengAnh_Nghe" : sheet['U39'].value,
                "2_TiengAnh_Noi" : sheet['AA39'].value,
                "2_TiengAnh_Doc" : sheet['AG39'].value,
                "2_TiengAnh_Viet" : sheet['AM39'].value,

                "2_NNK_Ten" : sheet['G40'].value,
                "2_NNK_Nghe" : sheet['U40'].value,
                "2_NNK_Noi" : sheet['AA40'].value,
                "2_NNK_Doc" : sheet['AG40'].value,
                "2_NNK_Viet" : sheet['AM40'].value,

                "2_ChungChi_Toeic" : sheet['H42'].value,
                "2_ChungChi_Toefl" : sheet['N42'].value,
                "2_ChungChi_IELTS" : sheet['T42'].value,
                "2_ChungChi_Khac" : sheet['Z42'].value,
                "2_ChungChi_NgayThi" :  vld(0,null_vld(sheet['AM42'].value)),


                "3_DonViCongTac_1" : sheet['B47'].value,
                "3_ChucVu_1" : sheet['L47'].value,
                "3_ThoiGian_Start_1" :  vld(1,null_vld(sheet['Q47'].value)),
                "3_ThoiGian_End_1" :  vld(1,null_vld(sheet['U47'].value)),
                "3_QuanLy_1" : sheet['Y47'].value,
                "3_CongViecPhuTrach_1" : sheet['AD47'].value,
                "3_LyDoThoiViec_1" : sheet['AM47'].value,

                "3_DonViCongTac_2" : sheet['B48'].value,
                "3_ChucVu_2" : sheet['L48'].value,
                "3_ThoiGian_Start_2" :  vld(1,null_vld(sheet['Q48'].value)),
                "3_ThoiGian_End_2" :  vld(1,null_vld(sheet['U48'].value)),
                "3_QuanLy_2" : sheet['Y48'].value,
                "3_CongViecPhuTrach_2" : sheet['AD48'].value,
                "3_LyDoThoiViec_2" : sheet['AM48'].value,

                "3_DonViCongTac_3" : sheet['B49'].value,
                "3_ChucVu_3" : sheet['L49'].value,
                "3_ThoiGian_Start_3" :  vld(1,null_vld(sheet['Q49'].value)),
                "3_ThoiGian_End_3" :   vld(1,null_vld(sheet['U49'].value)),
                "3_QuanLy_3" : sheet['Y49'].value,
                "3_CongViecPhuTrach_3" : sheet['AD49'].value,
                "3_LyDoThoiViec_3" : sheet['AM49'].value,

                "3_DonViCongTac_4" : sheet['B50'].value,
                "3_ChucVu_4" : sheet['L50'].value,
                "3_ThoiGian_Start_4" :  vld(1,null_vld(sheet['Q50'].value)),
                "3_ThoiGian_End_4" :  vld(1,null_vld(sheet['U50'].value)),
                "3_QuanLy_4" : sheet['Y50'].value,
                "3_CongViecPhuTrach_4" : sheet['AD50'].value,
                "3_LyDoThoiViec_4" : sheet['AM50'].value,
                "3_ThoiGianConLai_Start_1" :  vld(0,null),
                "3_ThoiGianConLai_Start_2" :  vld(0,null),
                "3_ThoiGianConLai_Start_3" :  vld(0,null),
                "3_ThoiGianConLai_End_1" : vld(0,null),
                "3_ThoiGianConLai_End_2" : vld(0,null),
                "3_ThoiGianConLai_End_3" : vld(0,null),
                 


                "3_ThanhTichNoiBat" : sheet['B52'].value,
                "3_ThoiGianLamViecNganHang" : sheet['X52'].value,
                "3_ThoiGianLamViecKhac" : sheet['AG52'].value,
                "3_MucLuongGanNhat" : sheet['AK52'].value,


                "4_HovaTen_1" : sheet['B56'].value,
                "4_QuanHe_1" : sheet['L56'].value,
                "4_NamSinh_1" : vld(2,null_vld(sheet['Q56'].value)),
                "4_GioiTinh_1" : sheet['V56'].value,
                "4_NgheNghiep_NoiLV_1" : sheet['AA56'].value,
                "4_HovaTen_2" : sheet['B57'].value,
                "4_QuanHe_2" : sheet['L57'].value,
                "4_NamSinh_2" : vld(2,null_vld(sheet['Q57'].value)),
                "4_GioiTinh_2" : sheet['V57'].value,
                "4_NgheNghiep_NoiLV_2" : sheet['AA57'].value,
                "4_HovaTen_3" : sheet['B58'].value,
                "4_QuanHe_3" : sheet['L58'].value,
                "4_NamSinh_3" : vld(2,null_vld(sheet['Q58'].value)),
                "4_GioiTinh_3" : sheet['V58'].value,
                "4_NgheNghiep_NoiLV_3" : sheet['AA58'].value,
                "4_HovaTen_4" : sheet['B59'].value,
                "4_QuanHe_4" : sheet['L59'].value,
                "4_NamSinh_4" : vld(2,null_vld(sheet['Q59'].value)),
                "4_GioiTinh_4" : sheet['V59'].value,
                "4_NgheNghiep_NoiLV_4" :sheet['AA59'].value, 
                "4_HovaTen_5" : sheet['B60'].value,
                "4_QuanHe_5" : sheet['L60'].value,
                "4_NamSinh_5" : vld(2,null_vld(sheet['Q60'].value)),
                "4_GioiTinh_5" : sheet['V60'].value,
                "4_NgheNghiep_NoiLV_5" : sheet['AA60'].value,


                "5_PhamChatNangKhieu" : sheet['B64'].value,
                "5_KyNangDacBiet" : sheet['L64'].value,
                "5_ThanhVien" : sheet['V64'].value,
                "5_KhoaHoc" : sheet['AG64'].value,

                "6_DanhHieu" : sheet['B69'].value, 
                "6_Thoigian_DanhHieu" :  vld(0,null_vld(sheet['Q69'].value)),
                "6_HinhThucKyLuat" : sheet['X69'].value,
                "6_ThoiGianKyLuat" :  vld(0,null_vld(sheet['AM69'].value)),

                "7_newspaper" :  newspaper,
                "7_newssite" :  newssite,
                "7_website" :  website,
                "7_fanpage" :   fanpage,
                "7_socialmedia" :   social,
                "7_BanBeNguoiThan" :  relative,
                "7_MB" :  MBemp,
                "7_Khac" : sheet['AD73'].value,

                "8_HovaTen_1" :  sheet['B77'].value,
                "8_ChucVu_1" : sheet['L77'].value,
                "8_DonViCongTac_1" : sheet['Q77'].value,
                "8_MoiQuanHe_1" : sheet['Y77'].value,
                "8_Mobile_1" :  mobile_8_1,
                "8_email_1" :  email_8_1,

                "8_HovaTen_2" : sheet['B78'].value,
                "8_ChucVu_2" : sheet['L78'].value,
                "8_DonViCongTac_2" : sheet['Q78'].value,
                "8_MoiQuanHe_2" : sheet['Y78'].value,
                "8_Mobile_2" :  mobile_8_2,
                "8_email_2" :  email_8_2,

                "8_HovaTen_3" : sheet['B79'].value,
                "8_ChucVu_3" : sheet['L79'].value,
                "8_DonViCongTac_3" : sheet['Q79'].value,
                "8_MoiQuanHe_3" : sheet['Y79'].value,
                "8_Mobile_3" :   mobile_8_3,
                "8_email_3" :  email_8_3,

                "9_HovaTen_1" : sheet['B83'].value,
                "9_ChucVu_1" : sheet['L83'].value,
                "9_DonViCongTac_1" : sheet['Q83'].value,
                "9_MoiQuanHe_1" : sheet['Y83'].value,
                "9_Mobile_1" :  mobile_9_1,
                "9_email_1" :   email_9_1,
                "9_HovaTen_2" : sheet['B84'].value,
                "9_ChucVu_2" : sheet['L84'].value,
                "9_DonViCongTac_2" : sheet['Q84'].value,
                "9_MoiQuanHe_2" : sheet['Y84'].value,
                "9_Mobile_2" :  mobile_9_2,
                "9_email_2" :   email_9_2,
                "9_HovaTen_3" : sheet['B85'].value,
                "9_ChucVu_3" : sheet['L85'].value,
                "9_DonViCongTac_3" : sheet['Q85'].value,
                "9_MoiQuanHe_3" : sheet['Y85'].value,
                "9_Mobile_3" :  mobile_9_3,
                "9_email_3" :  email_9_3,

                "10_ViTriTungUngTuyen" : sheet['B89'].value,
                "10_ThoiGian" :  vld(0,null_vld(sheet['L89'].value)),
                "10_KhiemKhuyetNgoaiHinh_Co" : sheet['AG89'].value,
                "10_KhiemKhuyetNgoaiHinh_Khong" : sheet['AL89'].value,

                "10_DiemManh" : sheet['B91'].value,
                "10_DiemYeu" : sheet['V91'].value,

                "10_DienGiaiLyDo" : sheet['B93'].value,
                "10_NgayTao" :  vld(0,null_vld(sheet['D97'].value))
                        }
                        print('Sending data to Dataverse')
                        requests.post(cvupload, json = data)
                if cv_type == 'new':
                        print('Converting data to JSON')
                        if sheet['N83'].value == "Website tuyển dụng của MB": website = "1"
                        elif sheet['N83'].value == "Fanpage Tuyển dụng MBBank": fanpage = "2"
                        elif sheet['N83'].value == "Cán bộ nhân viên MB giới thiệu": MBemp = "5"
                        elif sheet['N83'].value == "Bạn bè/Người thân": relative = "4"
                        elif sheet['N83'].value == "LinkedIn": linkedin = "8"
                        elif sheet['N83'].value == "Vietnamwork": vnwork = "9"
                        elif sheet['N83'].value == "TopCV": topcv = "10"
                        elif sheet['N83'].value == "TopDev": topdev = "11"
                        elif sheet['N83'].value == "ITViec":itviec = "12"
                        elif sheet['N83'].value == "YBox": ybox = "13"
                        elif sheet['N83'].value == "CareerBuilder": careerbuilder = "14"
                        
                        
                               
                        data = {"owner": null_vld(owner),
                                "CV_image": CV_image,
                                "username": "admin",
                                "password": "@MB2023",
                                "CV_layout": "new2023",
                                "profile_picture": pp,
                                "NV1" :null_vld(sheet['B13'].value),
                                "NV2" : null_vld(sheet['P13'].value),
                                "DiaDiemLamViec" : null_vld(sheet['AJ13'].value),

                                "1_HovaTen" : null_vld(sheet['B18'].value.title().strip()),
                                "1_NgaySinh" :  vld(0,null_vld(sheet['Q18'].value) +"/"+ null_vld(sheet['S18'].value)+"/"+null_vld(sheet['U18'].value)),
                                "1_GioiTinh" : null_vld(sheet['Y18'].value),
                                "1_ChieuCao" : null_vld(sheet['AG18'].value),
                                "1_CanNang" : null_vld(sheet['AM18'].value),

                                "1_thuongtru_tinh" : null_vld(sheet['B20'].value),
                                "1_thuongtru_quan" : null_vld(sheet['J20'].value),
                                "1_thuongtru_diachi" : null_vld(sheet['R20'].value),
                                "1_cell_personal" : null_vld(sheet['AG20'].value),

                                "1_NoiO_tinh" : null_vld(sheet['B22'].value),
                                "1_NoiO_quan" : null_vld(sheet['J22'].value),
                                "1_NoiO_diachi" : null_vld(sheet['R22'].value),
                                "1_TtHonNhan" : null_vld(sheet['AG22'].value),

                                "1_Email" : null_vld(sheet['B24'].value),
                                "1_Cmnd_passport" : null_vld(sheet['Q24'].value),
                                "1_NoiCap" : null_vld(sheet['Y24'].value),
                                "1_NgayCap" :  vld(0,null_vld(sheet['AG24'].value) +"/"+ null_vld(sheet['AH24'].value)+"/"+null_vld(sheet['AI24'].value)),
                                "1_DanToc" : null_vld(sheet['AM24'].value),

                                "1_UrgentTen" : null_vld(sheet['B26'].value),
                                "1_UrgentSdt" :  null_vld(sheet['AG26'].value),
                                "1_UrgentQuanHe" : null_vld(sheet['AM26'].value),
                                "1_NgayBatDau" :  null_vld(sheet['B28'].value),
                                "1_ThoiDiemLienHe" : null_vld(sheet['M28'].value),
                                "1_MucLuongDeNghi" : null_vld(sheet['AG28'].value),

                                "2_Ten_THPT" : null_vld(sheet['B33'].value),
                                "2_ChuyenNganh_THPT" : null_vld(sheet['Q33'].value),
                                "2_Thoigian_start_THPT" : vld(2,null_vld(sheet['Y33'].value)),
                                "2_Thoigian_end_THPT" : vld(1,null_vld(sheet['AC33'].value)+ "/"+null_vld(sheet['AD33'].value)),
                                "2_DiemTrungBinh_THPT" : null_vld(sheet['AG33'].value),
                                "2_LoaiHinhDaoTao_THPT" : null_vld(sheet['AP33'].value),

                                "2_Ten_TC" : null_vld(sheet['B34'].value),
                                "2_ChuyenNganh_TC" : null_vld(sheet['Q34'].value),
                                "2_Thoigian_start_TC" : vld(2,null_vld(sheet['Y34'].value)),
                                "2_Thoigian_end_TC" : vld(1,null_vld(sheet['AC34'].value)+ "/"+null_vld(sheet['AD34'].value)),
                                "2_DiemTrungBinh_TC" : null_vld(sheet['AG34'].value),
                                "2_LoaiHinhDaoTao_TC" : null_vld(sheet['AP34'].value),

                                "2_Ten_CD" : null_vld(sheet['B35'].value),
                                "2_ChuyenNganh_CD" : null_vld(sheet['Q35'].value),
                                "2_Thoigian_start_CD" : vld(2,null_vld(sheet['Y35'].value)),
                                "2_Thoigian_end_CD" : vld(1,null_vld(sheet['AC35'].value)+ "/"+null_vld(sheet['AD35'].value)),
                                "2_DiemTrungBinh_CD" : null_vld(sheet['AG35'].value),
                                "2_LoaiHinhDaoTao_CD" : null_vld(sheet['AP35'].value),

                                "2_Ten_DH" : null_vld(sheet['B36'].value),
                                "2_ChuyenNganh_DH" : null_vld(sheet['Q36'].value),
                                "2_Thoigian_start_DH" : vld(2,null_vld(sheet['Y36'].value)),
                                "2_Thoigian_end_DH" : vld(1,null_vld(sheet['AC36'].value)+ "/"+null_vld(sheet['AD36'].value)),
                                "2_DiemTrungBinh_DH" : null_vld(sheet['AG36'].value),
                                "2_LoaiHinhDaoTao_DH" : null_vld(sheet['AP36'].value),

                                "2_Ten_TDH" : null_vld(sheet['B37'].value),
                                "2_ChuyenNganh_TDH" : null_vld(sheet['Q37'].value),
                                "2_Thoigian_start_TDH" : vld(2,null_vld(sheet['Y37'].value)),
                                "2_Thoigian_end_TDH" : vld(1,null_vld(sheet['AC37'].value)+ "/"+null_vld(sheet['AD37'].value)),
                                "2_DiemTrungBinh_TDH" : null_vld(sheet['AG37'].value),
                                "2_LoaiHinhDaoTao_TDH" : null_vld(sheet['AP37'].value),

                                "2_TiengAnh_Nghe" : null_vld(sheet['U39'].value),
                                "2_TiengAnh_Noi" : null_vld(sheet['AA39'].value),
                                "2_TiengAnh_Doc" : null_vld(sheet['AG39'].value),
                                "2_TiengAnh_Viet" : null_vld(sheet['AM39'].value),

                                "2_NNK_Ten" : null_vld(sheet['G41'].value),
                                "2_NNK_Nghe" : null_vld(sheet['U41'].value),
                                "2_NNK_Noi" : null_vld(sheet['AA41'].value),
                                "2_NNK_Doc" : null_vld(sheet['AG41'].value),
                                "2_NNK_Viet" : null_vld(sheet['AM41'].value),

                                "2_ChungChi_Toeic" : null_vld(sheet['H43'].value),
                                "2_ChungChi_Toefl" : null_vld(sheet['N43'].value),
                                "2_ChungChi_IELTS" : null_vld(sheet['T43'].value),
                                "2_ChungChi_Khac" : null_vld(sheet['Z43'].value),
                                "2_ChungChi_NgayThi" :  vld(0,null_vld(sheet['AM43'].value) +"/"+ null_vld(sheet['AN43'].value)+"/"+null_vld(sheet['AO43'].value)),


                                "2_Giaithuong" : null_vld(sheet['B45'].value),
                                "2_ThongTinThemVeGiaiThuong" : null_vld(sheet['I45'].value),
                                "2_ThongTinKhac" : null_vld(sheet['T45'].value),
                                "2_VaiTroToChucXaHoi" : null_vld(sheet['AI45'].value),


                                "3_DonViCongTac_1" : null_vld(sheet['B51'].value),
                                "3_ChucVu_1" : null_vld(sheet['L51'].value),
                                "3_ThoiGian_Start_1" :  vld(1,null_vld(sheet['Q51'].value) +"/" +null_vld(sheet['R51'].value)),
                                "3_ThoiGian_End_1" :  vld(1,null_vld(sheet['U51'].value) +"/" +null_vld(sheet['V51'].value)),
                                "3_QuanLy_1" : null_vld(sheet['Y51'].value),
                                "3_CongViecPhuTrach_1" : null_vld(sheet['AD51'].value),
                                "3_LyDoThoiViec_1" : null_vld(sheet['AM51'].value),

                                "3_DonViCongTac_2" : null_vld(sheet['B52'].value),
                                "3_ChucVu_2" : null_vld(sheet['L52'].value),
                                "3_ThoiGian_Start_2" :  vld(1,null_vld(sheet['Q52'].value) +"/" +null_vld(sheet['R52'].value)),
                                "3_ThoiGian_End_2" :  vld(1,null_vld(sheet['U52'].value) +"/" +null_vld(sheet['V52'].value)),
                                "3_QuanLy_2" : null_vld(sheet['Y52'].value),
                                "3_CongViecPhuTrach_2" : null_vld(sheet['AD52'].value),
                                "3_LyDoThoiViec_2" : null_vld(sheet['AM52'].value),

                                "3_DonViCongTac_3" : null_vld(sheet['B53'].value),
                                "3_ChucVu_3" : null_vld(sheet['L53'].value),
                                "3_ThoiGian_Start_3" :  vld(1,null_vld(sheet['Q53'].value)  +"/" +null_vld(sheet['R53'].value)),
                                "3_ThoiGian_End_3" :  vld(1,null_vld(sheet['U53'].value)  +"/" +null_vld(sheet['V53'].value)),
                                "3_QuanLy_3" : null_vld(sheet['Y53'].value),
                                "3_CongViecPhuTrach_3" : null_vld(sheet['AD53'].value),
                                "3_LyDoThoiViec_3" : null_vld(sheet['AM53'].value),

                                "3_ThoiGian_Start_4" :  vld(0,null),
                                "3_ThoiGian_End_4" :  vld(0,null),


                                "3_DonViCongTacConLai_1" : null_vld(sheet['B57'].value),
                                "3_ChucVuConLai_1" : null_vld(sheet['L57'].value),
                                "3_ThoiGianConLai_Start_1" :  vld(1,null_vld(sheet['Q57'].value)  +"/"+null_vld(sheet['R57'].value)),
                                "3_ThoiGianConLai_End_1" :      vld(1,null_vld(sheet['U57'].value)  +"/"+null_vld(sheet['V57'].value)),                           
                                "3_CongViecPhuTrachConLai_1" : null_vld(sheet['Y57'].value),

                                "3_DonViCongTacConLai_2" : null_vld(sheet['B58'].value),
                                "3_ChucVuConLai_2" : null_vld(sheet['L58'].value),
                                "3_ThoiGianConLai_Start_2" :  vld(1,null_vld(sheet['Q58'].value)  +"/"+ null_vld(sheet['R58'].value)),
                                "3_ThoiGianConLai_End_2" :  vld(1,null_vld(sheet['U58'].value)  +"/"+null_vld(sheet['V58'].value)),
                                "3_CongViecPhuTrachConLai_2" : null_vld(sheet['Y58'].value),

                                "3_DonViCongTacConLai_3" : null_vld(sheet['B59'].value),
                                "3_ChucVuConLai_3" : null_vld(sheet['L59'].value),
                                "3_ThoiGianConLai_Start_3" :  vld(1,null_vld(sheet['Q59'].value)  +"/"+null_vld(sheet['R59'].value)),
                                "3_ThoiGianConLai_End_3" : vld(1,null_vld(sheet['U59'].value)  +"/"+null_vld(sheet['V59'].value)),                                
                                "3_CongViecPhuTrachConLai_3" : null_vld(sheet['Y59'].value),

                                "3_ThanhTichNoiBat" : null_vld(sheet['B61'].value),
                                "3_MucLuongGanNhat" : null_vld(sheet['AK61'].value),


                                "4_HovaTen_1" : null_vld(sheet['B65'].value),
                                "4_QuanHe_1" : null_vld(sheet['L65'].value),
                                "4_NamSinh_1" :null_vld(sheet['Q65'].value), 
                                "4_GioiTinh_1" : null_vld(sheet['V65'].value),
                                "4_NgheNghiep_NoiLV_1" : null_vld(sheet['AA65'].value),

                                "4_HovaTen_2" : null_vld(sheet['B66'].value),
                                "4_QuanHe_2" : null_vld(sheet['L66'].value),
                                "4_NamSinh_2" : null_vld(sheet['Q66'].value),
                                "4_GioiTinh_2" : null_vld(sheet['V66'].value),
                                "4_NgheNghiep_NoiLV_2" : null_vld(sheet['AA66'].value),

                                "4_HovaTen_3" : null_vld(sheet['B67'].value),
                                "4_QuanHe_3" : null_vld(sheet['L67'].value),
                                "4_NamSinh_3" : null_vld(sheet['Q67'].value),
                                "4_GioiTinh_3" : null_vld(sheet['V67'].value),
                                "4_NgheNghiep_NoiLV_3" : null_vld(sheet['AA67'].value),

                                "4_HovaTen_4" : null_vld(sheet['B68'].value),
                                "4_QuanHe_4" : null_vld(sheet['L68'].value),
                                "4_NamSinh_4" : null_vld(sheet['Q68'].value),
                                "4_GioiTinh_4" : null_vld(sheet['V68'].value),
                                "4_NgheNghiep_NoiLV_4" : null_vld(sheet['AA68'].value),
                                "4_HovaTen_5" : null_vld(sheet['B69'].value),
                                "4_QuanHe_5" : null_vld(sheet['L69'].value),
                                "4_NamSinh_5" : null_vld(sheet['Q69'].value),
                                "4_GioiTinh_5" : null_vld(sheet['V69'].value),
                                "4_NgheNghiep_NoiLV_5" : null_vld(sheet['AA69'].value),
                                "4_HovaTen_6" : null_vld(sheet['B70'].value),
                                "4_QuanHe_6" : null_vld(sheet['L70'].value),
                                "4_NamSinh_6" : null_vld(sheet['Q70'].value),
                                "4_GioiTinh_6" :  null_vld(sheet['V70'].value),
                                "4_NgheNghiep_NoiLV_6" : null_vld(sheet['AA70'].value),

                                "5_PhamChatNangKhieu" : null_vld(sheet['B74'].value),
                                "5_KyNangDacBiet" : null_vld(sheet['L74'].value),
                                "5_ThanhVien" : null_vld(sheet['V74'].value),
                                "5_KhoaHoc" : null_vld(sheet['AG74'].value),

                                "6_DanhHieu" : null_vld(sheet['B79'].value),
                                "6_HinhThucKyLuat" : null_vld(sheet['X79'].value),
                                "6_ThoiGianKyLuat" :  vld(0,null_vld(sheet['AM79'].value) +"/"+ null_vld(sheet['AN79'].value)+"/"+null_vld(sheet['AO79'].value)),

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
                                "7_Khac" : null_vld(sheet['AE83'].value),

                                "8_HovaTen_1" : null_vld(sheet['B87'].value),
                                "8_ChucVu_1" : null_vld(sheet['L87'].value),
                                "8_DonViCongTac_1" : null_vld(sheet['Q87'].value),
                                "8_MoiQuanHe_1" : null_vld(sheet['Y87'].value),
                                "8_Mobile_1" : null_vld(sheet['AG87'].value),
                                "8_email_1" : null_vld(sheet['AM87'].value), 

                                "8_HovaTen_2" : null_vld(sheet['B88'].value),
                                "8_ChucVu_2" : null_vld(sheet['L88'].value),
                                "8_DonViCongTac_2" : null_vld(sheet['Q88'].value),
                                "8_MoiQuanHe_2" : null_vld(sheet['Y88'].value),
                                "8_Mobile_2" : null_vld(sheet['AG88'].value),
                                "8_email_2" : null_vld(sheet['AM88'].value),

                                "8_HovaTen_3" : null_vld(sheet['B89'].value),
                                "8_ChucVu_3" : null_vld(sheet['L89'].value),
                                "8_DonViCongTac_3" : null_vld(sheet['Q89'].value),
                                "8_MoiQuanHe_3" : null_vld(sheet['Y89'].value),
                                "8_Mobile_3" : null_vld(sheet['AG89'].value),
                                "8_email_3" : null_vld(sheet['AM89'].value),

                                "9_HovaTen_1" : null_vld(sheet['B93'].value),
                                "9_ChucVu_1" : null_vld(sheet['L93'].value),
                                "9_DonViCongTac_1" : null_vld(sheet['Q93'].value),
                                "9_MoiQuanHe_1" : null_vld(sheet['Y93'].value),
                                "9_Mobile_1" : null_vld(sheet['AG93'].value),
                                "9_email_1" : null_vld(sheet['AM93'].value),

                                "9_HovaTen_2" : null_vld(sheet['B94'].value),
                                "9_ChucVu_2" : null_vld(sheet['L94'].value),
                                "9_DonViCongTac_2" : null_vld(sheet['Q94'].value),
                                "9_MoiQuanHe_2" : null_vld(sheet['Y94'].value),
                                "9_Mobile_2" : null_vld(sheet['AG94'].value),
                                "9_email_2" : null_vld(sheet['AM94'].value),

                                "9_HovaTen_3" : null_vld(sheet['B95'].value),
                                "9_ChucVu_3" : null_vld(sheet['L95'].value),
                                "9_DonViCongTac_3" : null_vld(sheet['Q95'].value),
                                "9_MoiQuanHe_3" : null_vld(sheet['Y95'].value),
                                "9_Mobile_3" : null_vld(sheet['AG95'].value),
                                "9_email_3" : null_vld(sheet['AM95'].value),

                                "10_ViTriTungUngTuyen" : null_vld(sheet['B99'].value),
                                "10_ThoiGian" :  vld(0,null_vld(sheet['L99'].value) +"/"+ null_vld(sheet['N99'].value)+"/"+null_vld(sheet['P99'].value)),
                                "10_KhiemKhuyetNgoaiHinh_Co" : null_vld(sheet['AG99'].value),
                                "10_KhiemKhuyetNgoaiHinh_Khong" : null_vld(sheet['AL99'].value),

                                "10_DiemManh" : null_vld(sheet['B101'].value),
                                "10_DiemYeu" : null_vld(sheet['V101'].value),
                                "10_LyDo_1" : null_vld(sheet['D103'].value),
                                "10_LyDo_2" : null_vld(sheet['D104'].value),
                                "10_LyDo_3" : null_vld(sheet['D105'].value),
                                "10_DienGiaiLyDo" : null_vld(sheet['Q103'].value),
                                "10_KhaNangDacBiet": null_vld(sheet['Q107'].value),
                                "10_NgayTao" :  vld(0,null_vld(sheet['B110'].value) +"/"+ null_vld(sheet['D110'].value)+"/"+null_vld(sheet['F110'].value))      
                                }
                        print('Sending data to Dataverse')
                        # for key, value in data.items():
                        #         key_type = type(key).__name__
                        #         value_type = type(value).__name__
                        #         print(f"Key: {key},Value: {value}, Type: {value_type}")
                        requests.post(cvupload, json = data)
                        # print('xxxx')
                wb1.close()
                kill()
        #-------------------------------------------------------

                os.remove(r"C:\\MB\\CV\\"+a[0])
                
                print('DONE')

        except Exception as error1:
                if blank == False:
                        os.remove(r"C:\\MB\\CV\\"+a[0])
                        erroritem = {"filename": str(a[0]), "owner": str(owner)}
                        print('Moved to error folder')
                        print(error1)
                        # requests.post(errorurl, json = erroritem)  
                else:
                        raise ValueError('No CV here')

    except Exception as error:
        print(error)
        kill()
        print(datetime.now())
        if datetime.now().hour == 23 and datetime.now().minute == 59 and datetime.now().second > 55:
               kill()
               excel = win32.gencache.EnsureDispatch('Excel.Application')
        time.sleep(2)
        pass


