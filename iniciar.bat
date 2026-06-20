@echo off
title GlobalCopyPlot - Servidor
color 0B
echo ===================================
echo   Iniciando GlobalCopyPlot...
echo ===================================
echo.

:: Activar entorno virtual
call venv\Scripts\activate

:: Iniciar Flask
echo.
python app.py

:: Si hay un error, mantener la ventana abierta
pause