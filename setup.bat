@echo off
REM ============================================================
REM  Monitor Retail LATAM - instalacion (solo la primera vez)
REM  Crea el entorno virtual .venv e instala requirements.txt
REM ============================================================
setlocal
cd /d "%~dp0"
echo.
echo === Monitor Retail LATAM: preparando el entorno ===

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo [ERROR] No encontre Python en este computador.
  echo         Instalalo desde el Centro de Software de tu empresa o desde https://www.python.org/downloads/
  echo         marca la casilla "Add python.exe to PATH" y vuelve a ejecutar setup.bat
  pause
  exit /b 1
)
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)"
if errorlevel 1 (
  echo [ERROR] Se requiere Python 3.9 o superior.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creando entorno virtual .venv ...
  %PY% -m venv .venv
  if errorlevel 1 (
    echo [ERROR] No se pudo crear el entorno virtual.
    pause
    exit /b 1
  )
)

echo Instalando librerias de requirements.txt ...
".venv\Scripts\python.exe" -m pip install --upgrade pip --disable-pip-version-check -q
".venv\Scripts\python.exe" -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 (
  echo.
  echo Reintentando con los certificados de Windows - redes con inspeccion SSL ...
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt --use-feature=truststore --disable-pip-version-check
)
if errorlevel 1 (
  echo.
  echo [ERROR] No se pudieron instalar las librerias.
  echo         Si tu red usa proxy, define HTTPS_PROXY antes de correr setup.bat, por ejemplo:
  echo         set HTTPS_PROXY=http://usuario:clave@proxy.empresa.com:8080
  pause
  exit /b 1
)

if not exist ".env" (
  copy ".env.example" ".env" >nul
  echo Se creo el archivo .env: abrelo y pega tus tokens gratuitos de INEGI y Banxico.
)
echo.
echo === Listo. Ahora ejecuta run.bat (o en VS Code: Terminal ^> Run Task ^> Actualizar Monitor) ===
endlocal
exit /b 0
