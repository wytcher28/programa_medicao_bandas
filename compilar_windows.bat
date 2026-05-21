@echo off
REM Script para compilar o programa no Windows.
REM Execute este arquivo dentro da pasta do projeto.

python -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
pyinstaller --onefile --name MedidasBanda medidas_banda.py

echo.
echo Executavel gerado em: dist\MedidasBanda.exe
echo Copie dados_log_bruto.xlsx para a mesma pasta do executavel antes de executar.
pause
