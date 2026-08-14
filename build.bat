@echo off
REM 文档快速排版 —— 打包为单文件 Windows exe
REM 用法：双击本文件，或在项目目录运行 build.bat
setlocal
cd /d "%~dp0"

set PY=venv314\Scripts\pyinstaller.exe
if not exist "%PY%" (
    echo [错误] 未找到 venv314 中的 pyinstaller，请先运行：venv314\Scripts\python.exe -m pip install pyinstaller
    pause
    exit /b 1
)

"%PY%" --noconfirm --clean DocFormatter.spec

if exist "dist\DocFormatter.exe" (
    copy /Y "dist\DocFormatter.exe" ".\DocFormatter.exe" >nul
    echo.
    echo [完成] 已生成单文件 exe：DocFormatter.exe
) else (
    echo [失败] 未生成 exe，请检查上方报错。
)
pause
