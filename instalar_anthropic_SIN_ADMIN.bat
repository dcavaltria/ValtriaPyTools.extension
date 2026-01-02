@echo off
setlocal EnableExtensions EnableDelayedExpansion

echo.
echo ============================================
echo   Instalar anthropic (sin administrador)
echo   ValtriaPyTools - pyRevit
echo ============================================
echo.

rem Carpeta de destino donde el script de Claude busca la libreria
set "SITE_PACKAGES=%USERPROFILE%\.pyrevit\site-packages"

rem Intentar localizar el interprete de Python que trae pyRevit
set "PYREVIT_ROOT=C:\Program Files\pyRevit-Master"
set "PYENGINE_DIR=%PYREVIT_ROOT%\bin\cengines"
set "PYTHON_EXE="

if exist "%PYENGINE_DIR%" (
    for /f "delims=" %%I in ('dir /b /ad "%PYENGINE_DIR%" 2^>nul ^| sort /r') do (
        if exist "%PYENGINE_DIR%\%%I\python.exe" (
            set "PYTHON_EXE=%PYENGINE_DIR%\%%I\python.exe"
            goto :python_ready
        )
    )
)

:python_ready
if not defined PYTHON_EXE (
    echo [WARN] No se encontro Python interno de pyRevit. Buscando Python del sistema...
    for /f "delims=" %%P in ('where python 2^>nul') do (
        if exist "%%~fP" (
            set "PYTHON_EXE=%%~fP"
            goto :python_ready_2
        )
    )
)

:python_ready_2
if not defined PYTHON_EXE (
    echo [ERROR] No se encontro ningun interprete de Python.
    echo Instala Python 3.10+ desde https://www.python.org/downloads/
    goto :fail
)

echo [INFO] Usando Python: %PYTHON_EXE%
echo.

if not exist "%SITE_PACKAGES%" (
    echo [INFO] Creando carpeta destino: %SITE_PACKAGES%
    mkdir "%SITE_PACKAGES%" >nul 2>nul
)

rem Calcular rutas locales para instalar pip sin privilegios
set "PIP_INFO="
set "PY_TMP=%TEMP%\pyrevit_paths_%RANDOM%.txt"
"%PYTHON_EXE%" -c "import os, sys; base=os.path.join(os.path.expanduser('~'), '.pyrevit'); root=os.path.join(base, 'python{}{}'.format(sys.version_info.major, sys.version_info.minor)); pkg=os.path.join(root, 'Python{}{}'.format(sys.version_info.major, sys.version_info.minor), 'site-packages'); scripts=os.path.join(root, 'Python{}{}'.format(sys.version_info.major, sys.version_info.minor), 'Scripts'); print(pkg + '|' + scripts + '|' + root)" > "%PY_TMP%" 2>nul
set /p PIP_INFO=<"%PY_TMP%"
del "%PY_TMP%" >nul 2>nul
if not defined PIP_INFO (
    echo [ERROR] No se pudo determinar la ruta local para pip.
    goto :fail
)
for /f "tokens=1,2,3 delims=|" %%A in ("%PIP_INFO%") do (
    set "PIP_SITE=%%~A"
    set "PIP_SCRIPTS=%%~B"
    set "PIP_BASE=%%~C"
)

if not defined PIP_SITE (
    echo [ERROR] No se pudo determinar la ruta local para pip.
    goto :fail
)

for %%D in ("%PIP_SITE%" "%PIP_SCRIPTS%" "%PIP_BASE%") do (
    if not exist "%%~D" mkdir "%%~D" >nul 2>nul
)

set "GETPIP=%TEMP%\get-pip.py"

if exist "%PIP_SITE%\pip" (
    echo [INFO] Pip ya esta instalado en %PIP_SITE%
) else (
    echo [INFO] Instalando pip en %PIP_BASE%
    if not exist "%GETPIP%" (
        echo [INFO] Descargando get-pip.py...
        powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%GETPIP%'" || goto :fail
    )
    "%PYTHON_EXE%" -c "import os, runpy; os.environ['PYTHONUSERBASE']=r'%PIP_BASE%'; runpy.run_path(r'%GETPIP%', run_name='__main__')" || goto :fail
)

if exist "%SITE_PACKAGES%\anthropic" (
    echo [INFO] Se detecto anthropic previamente instalado en %SITE_PACKAGES%.
    goto :verify_install
)

echo [INFO] Instalando anthropic en %SITE_PACKAGES%
"%PYTHON_EXE%" -c "import sys; sys.path.insert(0, r'%PIP_SITE%'); from pip._internal.cli.main import main; raise SystemExit(main(['install','anthropic','--upgrade','--no-warn-script-location','--target', r'%SITE_PACKAGES%']))" || goto :fail

:verify_install
echo [INFO] Verificando instalacion...
"%PYTHON_EXE%" -c "import sys; sys.path.insert(0, r'%SITE_PACKAGES%'); import anthropic; print('anthropic version:', getattr(anthropic, '__version__', 'OK'))" || goto :fail

echo.
echo [SUCCESS] La libreria anthropic quedo instalada en:
echo          %SITE_PACKAGES%
echo.
echo Si Revit ya estaba abierto, recarga pyRevit para que reconozca la libreria.
goto :exit_ok

:fail
echo.
echo [ERROR] La instalacion no se completo.
echo Puedes intentar instalar manualmente ejecutando:
echo    python -m pip install anthropic --target="%SITE_PACKAGES%"
echo.
pause
exit /b 1

:exit_ok
echo.
pause
exit /b 0
