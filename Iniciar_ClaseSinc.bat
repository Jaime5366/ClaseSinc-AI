@echo off
title ClaseSinc AI - Iniciador de la Aplicacion
color 0B
echo =====================================================================
echo                 INICIANDO CLASESINC AI 🎓
echo =====================================================================
echo.
cd /d "%~dp0"
echo [1/2] Activando entorno virtual de Python (venv)...
call .\venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [ERROR] No se pudo activar el entorno virtual. Asegurate de que la carpeta 'venv' existe.
    pause
    exit /b
)
echo.
echo [2/2] Lanzando el servidor de Streamlit...
echo La aplicacion deberia abrirse automaticamente en tu navegador.
echo Si no lo hace, ingresa manualmente a: http://localhost:8501
echo.
echo Para cerrar la aplicacion, cierra esta ventana.
echo =====================================================================
streamlit run app.py
pause
