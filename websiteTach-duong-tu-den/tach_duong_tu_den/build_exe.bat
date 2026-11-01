@echo off
REM Dong goi thanh 1 file .exe (may dich khong can cai Python).
REM config.py bi loai ra de API key khong nam trong file .exe.
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name TachDuongTuDen --exclude-module config main.py
echo.
echo Xong. File nam o: dist\TachDuongTuDen.exe
pause
