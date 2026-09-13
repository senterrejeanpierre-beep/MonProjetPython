@echo off
pushd "%~dp0"
py -3 -m unittest test_creation_chantier test_stabilite_ciblee test_pilotage_avenants_moins test_bordereau_public test_partage_multiplateforme
set "horizon_result=%errorlevel%"
popd
pause
exit /b %horizon_result%
