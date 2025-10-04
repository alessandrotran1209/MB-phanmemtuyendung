from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
import uvicorn
import base64
import csv
import io
import re


app = FastAPI()

@app.post('/mbcv')
def convertcv(inputdata: dict):
    try:
        bytestring = inputdata.get("bytestring","")
        filename = inputdata.get("filename","")
        binary_data = base64.b64decode(bytestring)

        with open(r"C:\\MB\\CV\\"+filename, "wb") as file:
            file.write(binary_data)
        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/dxtd')
def dexuattuyendung(inputdata: dict):
    try:
        bytestring = inputdata.get("bytestring","")
        filename = inputdata.get("filename","")
        binary_data = base64.b64decode(bytestring)

        with open(r"C:\\MB\\DXTD\\"+filename, "wb") as file:
            file.write(binary_data)
        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/dxpvnb')
def dexuatphongvan(inputdata: dict):
    try:
        bytestring = inputdata.get("bytestring","")
        filename = inputdata.get("filename","")
        binary_data = base64.b64decode(bytestring)

        with open(r"C:\\MB\\DXPVNB\\"+filename, "wb") as file:
            file.write(binary_data)
        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")
    
@app.post('/sms')
def sms(smsdata: dict):
    try:
        CSVstring = smsdata.get("body","")
        filename = smsdata.get("filename","")
        lines = CSVstring.split('\n')

# Specify the file path and open the CSV file in write mode
        with open(r"C:\\MB\\SMS\\"+filename, mode='w', newline='',encoding='utf-8') as csv_file:
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
                    csv_writer.writerow(str(field) for field in row)

        return JSONResponse(content={"message": "DONE"}, status_code=200)
    except:
        raise HTTPException(status_code=500, detail="error")

@app.post('/hcm')
def exportHCM(smsdata: dict):
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

@app.post('/thkqtd')
def thkqtd(dsuv: dict):
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
def dxl(smsdata: dict):
    try:
        CSVstring = smsdata.get("body","")
        filename = smsdata.get("filename","")
        lines = [x[0] for x in re.findall(r'((?:(\"+)[\s\S]+?\2|[^"\n]+)+)', CSVstring)]

# Specify the file path and open the CSV file in write mode
        with open(r"C:\\MB\\Export_DXL\\"+filename, mode='w', newline='',encoding='utf-8') as csv_file:
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

@app.post('/dstv')
def exportDS_ThiViet(dstv: dict):
    try:
        CSVstring = dstv.get("body","")
        filename = dstv.get("filename","")
        lines = [x[0] for x in re.findall(r'((?:(\"+)[\s\S]+?\2|[^"\n]+)+)', CSVstring)]
        with open(r"C:\\MB\\Export_DSTV\\"+filename, mode='w', newline='',encoding='utf-8') as csv_file:
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

@app.post('/dsuv') #Theo mau Database.xlsm
def exportDSUV(dstv: dict):
    try:
        CSVstring = dstv.get("body","")
        filename = dstv.get("filename","")
        lines = [x[0] for x in re.findall(r'((?:(\"+)[\s\S]+?\2|[^"\n]+)+)', CSVstring)]
        with open(r"C:\\MB\\DSUV\\"+filename, mode='w', newline='',encoding='utf-8') as csv_file:
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

if __name__ == '__main__':
    uvicorn.run(app, host='0.0.0.0', port=8080)