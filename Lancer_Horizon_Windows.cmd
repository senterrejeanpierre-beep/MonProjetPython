@echo off
pushd "%~dp0"
py -3 main.py
if errorlevel 1 pause
popd
