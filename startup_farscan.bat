@echo off
chcp 1251
set COMM="start"
:: ���� ������ ���� � python-�������
cd "C:\"Documents and Settings"\farscan\"������� ����""
py -u "farscan.py" %COMM%
echo %ERRORLEVEL%
pause