@echo off
echo Running Python Scripts...

start python C:\Users\phanmemtuyendung\API_2024.py
timeout /t 50 /nobreak >nul
start python C:\Users\phanmemtuyendung\Supervisor.py


echo All Python scripts started.
exit

