@echo off
setlocal
cd /d "%~dp0"
set "VENV=%~dp0..\.venv"

if not exist "%VENV%\Scripts\python.exe" (
    echo 正在创建虚拟环境...
    python -m venv "%VENV%"
    if errorlevel 1 goto :error
)

"%VENV%\Scripts\python.exe" -m pip install -r requirements-dev.txt
if errorlevel 1 goto :error

"%VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed --name "乙方宝招标信息爬取工具" --hidden-import yfb_bid_spider --hidden-import yfb_browser_auth --hidden-import websocket --exclude-module numpy yfb_spider_app.py
if errorlevel 1 goto :error

copy /Y "2026-乙方宝招标信息统计.xlsx" "dist\乙方宝招标信息爬取工具\2026-乙方宝招标信息统计.xlsx"
copy /Y "乙方宝爬虫使用说明.md" "dist\乙方宝招标信息爬取工具\乙方宝爬虫使用说明.md"
copy /Y "乙方宝爬虫使用说明.txt" "dist\乙方宝招标信息爬取工具\乙方宝爬虫使用说明.txt"
echo.
echo 打包完成后请查看 dist\乙方宝招标信息爬取工具 文件夹
pause
exit /b 0

:error
echo.
echo 打包失败，请检查上方错误信息。
pause
exit /b 1
