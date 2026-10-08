@echo off
rem 微信文件地图 启动器（只读索引工具）
cd /d "%~dp0"
netstat -ano | findstr :8012 | findstr LISTENING >nul 2>nul
if errorlevel 1 (
  start "" /b "C:\Users\fanruocheng\AppData\Local\Programs\Python\Python312\pythonw.exe" "server.py"
  timeout /t 2 /nobreak >nul
)
start "" "http://127.0.0.1:8012"
