@echo off
REM ============================================================
REM  Monitor Retail LATAM - actualizacion
REM  Uso: doble clic.  Opciones: run.bat --prueba, --sin-noticias, --offline
REM ============================================================
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  call setup.bat
  if errorlevel 1 exit /b 1
)
set "PYTHONPATH=%~dp0src"
set "PYTHONIOENCODING=utf-8"
".venv\Scripts\python.exe" -m monitor %*
set "RC=%ERRORLEVEL%"
echo.
echo Excel: excel\Monitor_Retail_LATAM.xlsx   (hoja Log_Actualizacion = que cambio y de donde salio)
echo Reportes y respaldos: carpeta output\
pause
exit /b %RC%
