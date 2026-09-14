@echo off
setlocal
cd /d C:\Users\Rais\PhpstormProjects\allods

if not exist logs mkdir logs

set PY=C:\Users\Rais\AppData\Local\Python\pythoncore-3.14-64\python.exe
set SCRIPT=src\allods_hpi_to_xlsx.py

set YYYY=%date:~-4%
set MM=%date:~3,2%
set DD=%date:~0,2%

if defined YADISK_TOKEN (
  %PY% %SCRIPT% --upload-to-yadisk >> logs\allods_%YYYY%-%MM%-%DD%.log 2>&1
) else (
  %PY% %SCRIPT% >> logs\allods_%YYYY%-%MM%-%DD%.log 2>&1
)
endlocal
