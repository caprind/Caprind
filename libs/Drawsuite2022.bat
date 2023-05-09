@echo off

cls
echo.

echo Registrando arquivo, por favor aguarde...
cd %WinDir%\syswow64
Copy C:\Drawsuite2022\DrawSuite2022.ocx c:\Windows\syswow64

regsvr32/s C:\Program Files (x86)\Caprind\libs\DrawSuite2022.ocx

echo.
echo DrawSuite 2022 devidamente instalado em seu sistema!
echo.