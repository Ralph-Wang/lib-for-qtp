@echo off
set pwd=%cd%
cd .\test
main.vbs
more test.log
cd %pwd%
