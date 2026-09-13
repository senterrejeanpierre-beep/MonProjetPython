@echo off
pushd "%~dp0"
py -3 -m PyInstaller --clean --noconfirm main.spec
if errorlevel 1 pause
popd
