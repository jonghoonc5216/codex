@echo off
setlocal
cd /d "%~dp0"
python "%~dp0pdf_to_jpg.py" %*
