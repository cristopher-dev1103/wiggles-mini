@echo off
TITLE WIGGLES MINI
CLS
color 0A

:: ==========================================================
::  WIGGLES MINI - Inventario para visitas a clientes
::  No instala nada. No requiere modulos. 100% portable.
:: ==========================================================

NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Solicitando permisos de Administrador...
    powershell -NoProfile -Command "Start-Process '%~dpnx0' -Verb RunAs"
    EXIT /B
)

cd /d "%~dp0"

IF NOT EXIST "WigglesMini.ps1" (
    color 0C
    echo ERROR: No se encuentra WigglesMini.ps1
    echo Copia todos los archivos en la misma carpeta.
    pause
    EXIT /B 1
)

IF NOT EXIST "WigglesCore.ps1" (
    color 0C
    echo ERROR: No se encuentra WigglesCore.ps1
    echo Copia todos los archivos en la misma carpeta.
    pause
    EXIT /B 1
)

powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "WigglesMini.ps1"

EXIT /B %ERRORLEVEL%
