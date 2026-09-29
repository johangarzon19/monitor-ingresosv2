@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo   Monitor de Ingresos - servidor local
echo   Abre:  http://localhost:8080
echo   Detener: Ctrl + C
echo.
where py >nul 2>nul
if %errorlevel%==0 (
  py -m http.server 8080
  goto fin
)
where python >nul 2>nul
if %errorlevel%==0 (
  python -m http.server 8080
  goto fin
)
where node >nul 2>nul
if %errorlevel%==0 (
  npx --yes serve -l 8080 .
  goto fin
)
echo.
echo   No encontre Python ni Node instalados.
echo   Instala Python desde https://www.python.org/downloads/
echo   (marca la casilla "Add Python to PATH") y vuelve a ejecutar este archivo.
echo.
pause
:fin
