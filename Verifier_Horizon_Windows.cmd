@echo off
pushd "%~dp0"
py -3 -m unittest discover -q
set "horizon_result=%errorlevel%"
popd
pause
exit /b %horizon_result%
