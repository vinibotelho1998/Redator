@echo off
title Redator - gerar instalador
echo.
echo   Gerando o instalador do Redator.
echo   O resultado sai em out\Redator-Setup.exe
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0packaging\build.ps1"
pause
