@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON_EXE%" (
    echo 未找到本机的 Python 环境，请查看“先读我.txt”中的安装和运行说明。
    if /i not "%~1"=="--no-pause" pause
    exit /b 1
)
"%PYTHON_EXE%" -X utf8 "%~dp0douyin_weekly.py"
set "RUN_STATUS=%ERRORLEVEL%"
echo.
if /i not "%~1"=="--no-pause" pause
exit /b %RUN_STATUS%