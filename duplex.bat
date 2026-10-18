@echo off
"D:\envs\pdfocr\Scripts\python.exe" "%~dp0duplex_print.py" %*
if errorlevel 1 pause
