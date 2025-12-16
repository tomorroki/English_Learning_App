@echo off
title Learning App EXE Creator v3.6
chcp 65001 > nul
echo ========================================
echo  Learning App EXE Creator v3.6
echo  Google Translation Support
echo ========================================
echo.
cd /d "%~dp0"
echo Working directory: %CD%
echo.

REM Check virtual environment
if exist ".venv\Scripts\python.exe" (
    echo [1/4] Virtual environment found
    echo [1/4] Activating virtual environment...
    call .venv\Scripts\activate.bat
    
    echo.
    echo [2/4] Installing required packages...
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install pyinstaller tkcalendar deep-translator
    
    echo.
    echo [3/4] Building EXE file...
    .venv\Scripts\python.exe -m PyInstaller ^
        --clean ^
        --onefile ^
        --windowed ^
        --name "LearningApp" ^
        --add-data "sound_correct.mp3;." ^
        --add-data "sound_incorrect.mp3;." ^
        --hidden-import=tkcalendar ^
        --hidden-import=deep_translator ^
        --hidden-import=deep_translator.google ^
        --noconfirm ^
        integrated_learning_app.py
    
) else (
    echo [1/4] No virtual environment found
    echo [1/4] Using system Python...
    
    echo.
    echo [2/4] Installing required packages...
    python -m pip install --upgrade pip
    python -m pip install pyinstaller tkcalendar deep-translator
    
    echo.
    echo [3/4] Building EXE file...
    python -m PyInstaller ^
        --clean ^
        --onefile ^
        --windowed ^
        --name "LearningApp" ^
        --add-data "sound_correct.mp3;." ^
        --add-data "sound_incorrect.mp3;." ^
        --hidden-import=tkcalendar ^
        --hidden-import=deep_translator ^
        --hidden-import=deep_translator.google ^
        --noconfirm ^
        integrated_learning_app.py
)

echo.
echo ========================================
echo [4/4] Build Result
echo ========================================

if exist "dist\LearningApp.exe" (
    echo.
    echo ✅ SUCCESS! EXE file created successfully!
    echo.
    echo Location: dist\LearningApp.exe
    
    REM Get file size
    for %%A in ("dist\LearningApp.exe") do (
        set size=%%~zA
        set /a sizeMB=%%~zA/1048576
        echo File size: %%~zA bytes (approx. !sizeMB! MB)
    )
    
    echo.
    echo ========================================
    echo Included Features:
    echo ========================================
    echo ✅ Word and Sentence Management
    echo ✅ Quiz System (Choice / Listening)
    echo ✅ Wrong Answer Review (3-strike system)
    echo ✅ Date Filter
    echo ✅ Google Translation (Free, unlimited)
    echo ✅ Sound Effects
    echo.
    echo ❌ DeepL Translation (requires: pip install deepl)
    echo ❌ TTS Voice (requires: pip install edge-tts)
    echo ========================================
    
    echo.
    echo Opening dist folder...
    explorer dist
    
) else (
    echo.
    echo ❌ FAILED! EXE file was not created.
    echo.
    echo ========================================
    echo Troubleshooting:
    echo ========================================
    
    REM Check what went wrong
    if exist ".venv\Scripts\python.exe" (
        echo Checking virtual environment packages...
        .venv\Scripts\python.exe -c "import PyInstaller; print('  ✅ PyInstaller')" 2>nul || echo "  ❌ PyInstaller not found"
        .venv\Scripts\python.exe -c "import tkcalendar; print('  ✅ tkcalendar')" 2>nul || echo "  ❌ tkcalendar not found"
        .venv\Scripts\python.exe -c "import deep_translator; print('  ✅ deep_translator')" 2>nul || echo "  ❌ deep_translator not found"
    ) else (
        echo Checking system packages...
        python -c "import PyInstaller; print('  ✅ PyInstaller')" 2>nul || echo "  ❌ PyInstaller not found"
        python -c "import tkcalendar; print('  ✅ tkcalendar')" 2>nul || echo "  ❌ tkcalendar not found"
        python -c "import deep_translator; print('  ✅ deep_translator')" 2>nul || echo "  ❌ deep_translator not found"
    )
    
    echo.
    echo Please check the error messages above.
)

echo.
pause