@echo off
REM ============================================================================
REM  Registrador BMS -> Google Sheet (doble clic, sin terminal).
REM  Abre el registrador en modo interactivo: una vez abierto, pulsa ENTER cada
REM  vez que quieras subir una toma de modulos a la hoja. Cierra con Ctrl+C.
REM  La 1a vez instala las dependencias solo (pyserial, requests).
REM ============================================================================
cd /d "%~dp0"
chcp 65001 >nul

REM Detecta el lanzador de Python disponible (py o python)
where py >nul 2>nul && (set "PY=py") || (set "PY=python")

REM Instala dependencias (idempotente: si ya estan, no hace nada)
%PY% -m pip install -r requirements_serial.txt --quiet --disable-pip-version-check

REM Arranca en modo interactivo (ENTER = subir una toma)
%PY% serial_to_sheet.py

echo.
echo (registrador cerrado)
pause
