@echo off
echo ============================================
echo  INSTALADOR - Exportador de Inventarios
echo ============================================
echo.

REM Verificar Python
echo [1/3] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH.
    echo Descargalo desde: https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)
echo OK - Python detectado.
echo.

REM Actualizar pip
echo [2/3] Actualizando pip, setuptools y wheel...
python -m pip install --upgrade pip setuptools wheel
echo.

REM Instalar dependencias
echo [3/3] Instalando dependencias...
echo Esto puede tardar unos minutos la primera vez...
echo.
pip install -r requirements.txt

echo.
echo ============================================
echo  INSTALACION COMPLETADA
echo ============================================
echo.
echo Para ejecutar la aplicacion:
echo   python app.py
echo.
echo Luego abre tu navegador en:
echo   http://localhost:5000
echo.
pause
