@echo off
title Gateway Bot

REM Vai pra pasta do script
cd /d "C:\Users\laporciuncula\Documents\python\ligar_gateway"

REM Executa com python
python gateway.py

REM Se der erro, mostra mensagem
if %errorlevel% neq 0 (
    echo.
    echo ERRO ao executar o script!
)

pause