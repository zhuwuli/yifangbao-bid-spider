@echo off
setlocal
cd /d "%~dp0"
set "VENV=%~dp0..\.venv"

if not exist "%VENV%\Scripts\python.exe" (
    echo 正在创建虚拟环境...
    python -m venv "%VENV%"
    if errorlevel 1 goto :error
    "%VENV%\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :error
)

"%VENV%\Scripts\python.exe" yfb_spider_app.py
exit /b %errorlevel%

:error
echo.
echo 环境创建或依赖安装失败。
pause
exit /b 1
