@echo off
chcp 1251
set COMM="start"
:: Ниже указан путь к python-скрипту
cd "C:\"Documents and Settings"\farscan\"Рабочий стол""
py -u "farscan.py" %COMM%
echo %ERRORLEVEL%
pause