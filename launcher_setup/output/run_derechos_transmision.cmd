@echo off
set "PROJECT_DIR=C:\Users\lunta\OneDrive\Desktop\C23\derechosTransmision"
pushd "%PROJECT_DIR%"
py src\main.py
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
