@echo off
setlocal
cd /d "%~dp0"

:: 1. Inicia o Bridge de E-mail de forma silenciosa (sem janela de terminal)
:: O pythonw.exe roda scripts em background no Windows.
where pythonw >nul 2>&1
if %errorlevel% == 0 (
    start /b pythonw.exe email_bridge.py
) else (
    start /min python.exe email_bridge.py
)

:: 2. Pequena pausa para garantir que o servidor subiu
timeout /t 2 /nobreak >nul

:: 3. Abre o Dashboard no Chrome
start chrome "%~dp0index.html"

exit
