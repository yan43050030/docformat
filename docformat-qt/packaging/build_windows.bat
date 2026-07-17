@echo off
REM Build standalone DocFormatPro.exe (requires: pip install pyinstaller)
cd /d "%~dp0\.."

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name DocFormatPro ^
    --collect-data docx ^
    --hidden-import scripts.formatter ^
    --hidden-import scripts.punctuation ^
    --hidden-import scripts.analyzer ^
    --hidden-import scripts.converter ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    main.py

echo.
echo Done: dist\DocFormatPro.exe
pause
