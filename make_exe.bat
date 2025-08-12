@echo off
title Learning App EXE Creator (Fixed)

echo ========================================
echo  Learning App EXE Creator (Fixed)
echo ========================================
echo.

cd /d "%~dp0"
echo Working directory: %CD%
echo.

REM Check virtual environment
if exist ".venv\Scripts\python.exe" (
    echo Virtual environment found
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
    
    echo Installing PyInstaller in virtual environment...
    .venv\Scripts\python.exe -m pip install pyinstaller
    
    echo.
    echo Creating EXE file using virtual environment...
    .venv\Scripts\python.exe -m PyInstaller --onefile --windowed --name "LearningApp" --add-data "sound_correct.mp3;." --add-data "sound_incorrect.mp3;." integrated_learning_app.py
    
) else (
    echo No virtual environment found, using system Python
    
    echo Installing PyInstaller in system...
    python -m pip install pyinstaller
    
    echo.
    echo Creating EXE file using system Python...
    python -m PyInstaller --onefile --windowed --name "LearningApp" --add-data "sound_correct.mp3;." --add-data "sound_incorrect.mp3;." integrated_learning_app.py
)

echo.
echo ========================================
echo           Result Check
echo ========================================

if exist "dist\LearningApp.exe" (
    echo.
    echo SUCCESS! EXE file created!
    echo Location: dist\LearningApp.exe
    echo.
    REM Get file size
    for %%A in ("dist\LearningApp.exe") do echo File size: %%~zA bytes
    echo.
    explorer dist
) else (
    echo.
    echo FAILED! EXE file was not created
    echo.
    echo Let's try troubleshooting...
    echo.
    
    REM Check what went wrong
    if exist ".venv\Scripts\python.exe" (
        echo Checking virtual environment PyInstaller...
        .venv\Scripts\python.exe -c "import PyInstaller; print('PyInstaller is available')" 2>nul || echo "PyInstaller not found in virtual environment"
    ) else (
        echo Checking system PyInstaller...
        python -c "import PyInstaller; print('PyInstaller is available')" 2>nul || echo "PyInstaller not found in system"
    )
)

echo.
pause