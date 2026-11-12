@echo off
cd /d %~dp0
echo ============================================
echo  AUTOMACAO ITR - preenchimento de declaracoes
echo ============================================
echo.
echo Antes de continuar:
echo  1. Planilha planilha_clientes.xlsx preenchida (aba Clientes)
echo  2. Programa ITR aberto e MAXIMIZADO
echo.
pause
python automacao_itr.py
echo.
pause
