@echo off
echo ==================================================
echo 🚀 INICIANDO PIPELINE DE DADOS - VITALMED
echo ==================================================

:: 1. Entra na pasta base
cd /d "C:\Users\laporciuncula\Documents\python\CANCELADOS"

echo [1/4] 📉 Processando CANCELADOS...
"C:\Users\laporciuncula\AppData\Local\Microsoft\WindowsApps\python3.13.exe" "CanceladosCLAUDE.py"

echo.
echo [2/4] 🔄 Processando REATIVADOS...
"C:\Users\laporciuncula\AppData\Local\Microsoft\WindowsApps\python3.13.exe" "reativados.py"

echo.
echo [3/4] 📉 Processando CANCELADOS FEIRA...
"C:\Users\laporciuncula\AppData\Local\Microsoft\WindowsApps\python3.13.exe" "canceladosFeira.py"

echo.
echo [4/4] 🔄 Processando REATIVADOS FEIRA...
"C:\Users\laporciuncula\AppData\Local\Microsoft\WindowsApps\python3.13.exe" "reativadosFeira.py"
echo.
echo ==================================================
echo ✅ TODOS OS RELATORIOS FORAM FINALIZADOS!
echo ==================================================
pause