@echo off
cd /d "%~dp0"
echo Compilando quiz.py com PyInstaller...

pyinstaller --onefile --windowed --add-data "dados-suporte-quiz.json;." --add-data "logo.png;." quiz.py

echo Build finalizado. Executável em dist\quiz.exe
pause
