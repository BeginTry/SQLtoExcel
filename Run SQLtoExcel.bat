TITLE SQLtoExcel
ECHO OFF
COLOR 0E
CLS

CD %~dp0

REM ECHO "%~dp0"
REM ECHO "%~dp0Addendum.xlsx"

REM NOTE! If /ScriptsFolder ends with a backslash, add a space character afterwards.
REM https://stackoverflow.com/questions/1291291/how-to-accept-command-line-args-ending-in-backslash#1291306

REM SQL Authentication
REM SQLtoExcel.exe /Server=.\SQL2017 /Database=master /Login=sa /Password=password /ScriptsFolder="%~dp0 " /ExcelFile="%~dp0Addendum.xlsx"

REM Windows Authentication (login/pwd omitted)
SQLtoExcel.exe /Server=.\SQL2017 /Database=master /ScriptsFolder="%~dp0 " /ExcelFile="%~dp0SQLtoExcel.xlsx"

