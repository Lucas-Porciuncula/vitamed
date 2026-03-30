@echo off
REM Ativa o ambiente virtual (se existir)
call "C:\Users\laporciuncula\Documents\python\metasPorEquipe\.venv\Scripts\activate.bat"

REM Executa o script Python
python "C:\Users\laporciuncula\Documents\python\metasPorEquipe\whatsapp2.py"

REM Mantém a janela aberta pra debug (opcional)
pause
