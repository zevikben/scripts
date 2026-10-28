@echo off
REM Usage: convert_osh_discont.bat <input.xlsx> [output.xls]
REM Output is always saved in c:\veka\datafiles
cd /d c:\veka\scr
if "%~1"=="" (
  echo Usage: convert_osh_discont.bat input.xlsx [output.xls]
  exit /b 1
)
python convert_osh_discont.py %*
