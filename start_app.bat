@echo off
chcp 65001 >nul
title 統合学習アプリ
cd /d "%~dp0"

echo ==========================================
echo    統合学習アプリ (単語・例文・クイズ)
echo ==========================================
echo.

REM 仮想環境の確認・作成
if not exist .venv (
    echo 初回起動：仮想環境を作成しています...
    python -m venv .venv
    if errorlevel 1 (
        echo Python の仮想環境作成に失敗しました。
        echo Python がインストールされているか確認してください。
        pause
        exit /b 1
    )
)

REM 仮想環境をアクティベート
call .venv\Scripts\activate.bat

REM 必要なライブラリがあるかチェック
python -c "import tkinter" 2>nul
if errorlevel 1 (
    echo 必要なライブラリをインストールしています...
    pip install deepl edge-tts tkcalendar deep-translator
)

REM アプリを起動
echo アプリを起動中...
pythonw integrated_learning_app.py

REM エラーハンドリング
if errorlevel 1 (
    echo.
    echo アプリの起動に問題が発生しました。
    echo デバッグ情報を表示するには debug_start.bat を使用してください。
    pause
)