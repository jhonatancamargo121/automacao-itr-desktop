@echo off
cd /d %~dp0
echo Instalando dependencias e gerando o executavel...
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --onefile --name AutomacaoITR automacao_itr.py
echo.
echo Pronto: dist\AutomacaoITR.exe
echo Copie o .exe para a pasta da planilha (planilha_clientes.xlsx e config.json).
pause
