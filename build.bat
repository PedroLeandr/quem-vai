@echo off
echo Instalando dependencias...
pip install -r requirements.txt
pip install pyinstaller

echo Compilando...
pyinstaller --onefile --windowed --name quem-vai quem-vai.py

echo.
echo Pronto! O executavel esta em: dist\quem-vai.exe
pause
