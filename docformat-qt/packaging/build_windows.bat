@echo off
REM Build DocFormatPro-v{version}.exe (requires: pip install pyinstaller)
cd /d "%~dp0\.."

REM Extract version from source
for /f %%v in ('python -c "import re; m=re.search(r"VERSION = '(.+?)'", open('app/main_window.py',encoding='utf-8').read()); print(m.group(1) if m else 'unknown')"') do set VER=%%v

echo Building DocFormatPro-v%VER%.exe ...

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "DocFormatPro-v%VER%" ^
    --collect-data docx ^
    --icon assets/icon.ico ^
    --add-data "assets;assets" ^
    --add-data "templates;templates" ^
    --hidden-import scripts.formatter ^
    --hidden-import scripts.punctuation ^
    --hidden-import scripts.analyzer ^
    --hidden-import scripts.converter ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    main.py

echo.
echo Done: dist\DocFormatPro-v%VER%.exe
pause
