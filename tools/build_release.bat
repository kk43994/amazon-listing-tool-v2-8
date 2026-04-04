@echo off
setlocal
cd /d "%~dp0\.."

if not exist ".venv\Scripts\python.exe" (
  echo 未找到 .venv\Scripts\python.exe
  echo 请先在 Windows 上创建虚拟环境并安装依赖
  exit /b 1
)

".venv\Scripts\python.exe" tools\build_release.py

