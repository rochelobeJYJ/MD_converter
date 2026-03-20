@echo off
setlocal

:: Define Variables
set ANACONDA_BASE_DIR=C:\Users\user\anaconda3
set ENV_NAME=pdf2md_build_env
set PYTHON_VER=3.10

:: Direct Execution Paths
set ENV_PYTHON=%ANACONDA_BASE_DIR%\envs\%ENV_NAME%\python.exe

echo ===================================================
echo [0/5] Initializing Anaconda...
echo ===================================================
call "%ANACONDA_BASE_DIR%\Scripts\activate.bat" base

echo ===================================================
echo [1/5] Cleaning up old conda environment if exists...
echo ===================================================
call conda env remove -n %ENV_NAME% -y

echo ===================================================
echo [2/5] Creating pure conda environment '%ENV_NAME%'...
echo ===================================================
call conda create -n %ENV_NAME% python=%PYTHON_VER% -y

echo ===================================================
echo [3/5] Installing dependencies...
echo ===================================================
call "%ENV_PYTHON%" -m pip install -r requirements.txt
call "%ENV_PYTHON%" -m pip install pyinstaller

echo ===================================================
echo [4/5] Running build_helper.py for PyInstaller...
echo ===================================================
call "%ENV_PYTHON%" build_helper.py
if %ERRORLEVEL% neq 0 (
    echo ===================================================
    echo [ERROR] Build failed! Please check the logs above.
    echo ===================================================
    call conda env remove -n %ENV_NAME% -y
    pause
    exit /b %ERRORLEVEL%
)

echo ===================================================
echo [5/5] Final cleanup: Removing temporary environment...
echo ===================================================
call conda env remove -n %ENV_NAME% -y

echo ===================================================
echo Build pipeline finished successfully! ZIP created.
echo ===================================================
pause
