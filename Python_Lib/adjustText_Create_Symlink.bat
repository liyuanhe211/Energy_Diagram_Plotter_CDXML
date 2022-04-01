pushd %~dp0
powershell -Command "New-Item -ItemType SymbolicLink -Path .\adjustText.py -Target ..\..\Python_Lib\adjustText.py"
pause