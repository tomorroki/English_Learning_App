@echo off
title Create Distribution Package

echo ========================================
echo  Creating Distribution Package
echo ========================================
echo.

REM Create distribution folder
set DIST_FOLDER=LearningApp_Distribution
if exist "%DIST_FOLDER%" (
    echo Removing existing distribution folder...
    rmdir /s /q "%DIST_FOLDER%"
)
mkdir "%DIST_FOLDER%"

REM Copy exe file
if exist "dist\LearningApp.exe" (
    copy "dist\LearningApp.exe" "%DIST_FOLDER%\"
    echo OK: Copied LearningApp.exe
) else (
    echo ERROR: LearningApp.exe not found
    pause
    exit
)

REM Copy sound files if they exist
if exist "sound_correct.mp3" (
    copy "sound_correct.mp3" "%DIST_FOLDER%\"
    echo OK: Copied sound_correct.mp3
)
if exist "sound_incorrect.mp3" (
    copy "sound_incorrect.mp3" "%DIST_FOLDER%\"
    echo OK: Copied sound_incorrect.mp3
)

REM Create README file
echo Learning App > "%DIST_FOLDER%\README.txt"
echo ============ >> "%DIST_FOLDER%\README.txt"
echo. >> "%DIST_FOLDER%\README.txt"
echo HOW TO USE: >> "%DIST_FOLDER%\README.txt"
echo Just double-click "LearningApp.exe" >> "%DIST_FOLDER%\README.txt"
echo. >> "%DIST_FOLDER%\README.txt"
echo FEATURES: >> "%DIST_FOLDER%\README.txt"
echo - Word and sentence management >> "%DIST_FOLDER%\README.txt"
echo - Translation support (DeepL/Google) >> "%DIST_FOLDER%\README.txt"
echo - Multiple quiz types >> "%DIST_FOLDER%\README.txt"
echo - Date filtering >> "%DIST_FOLDER%\README.txt"
echo - Review system for wrong answers >> "%DIST_FOLDER%\README.txt"
echo - Text-to-speech functionality >> "%DIST_FOLDER%\README.txt"
echo. >> "%DIST_FOLDER%\README.txt"
echo NOTES: >> "%DIST_FOLDER%\README.txt"
echo - First startup may take a few seconds >> "%DIST_FOLDER%\README.txt"
echo - DeepL requires API key (Google Translate is free) >> "%DIST_FOLDER%\README.txt"

REM Show file size
for %%A in ("%DIST_FOLDER%\LearningApp.exe") do (
    set size=%%~zA
)
echo.
echo File size: %size% bytes

if %size% LSS 52428800 (
    echo Status: Excellent size (under 50MB)
) else if %size% LSS 104857600 (
    echo Status: Good size (under 100MB)
) else (
    echo Status: Large size (over 100MB, but still acceptable)
)

echo.
echo ========================================
echo  Distribution Package Complete!
echo ========================================
echo.
echo Package location: %DIST_FOLDER%
echo.
echo Contents:
dir /b "%DIST_FOLDER%"
echo.
echo TO DISTRIBUTE:
echo 1. Right-click the "%DIST_FOLDER%" folder
echo 2. Send to > Compressed (zipped) folder
echo 3. Share the ZIP file with others
echo.

explorer "%DIST_FOLDER%"
pause