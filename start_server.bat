@echo off
chcp 65001 > nul
title 受託確認票 業務管理サーバー
cd /d "%~dp0server"

if not exist "target\release\server.exe" (
    echo [ビルド] 初回起動のため release ビルドを実行します...
    cargo build --release
)

echo [起動] サーバーを開始します...
target\release\server.exe
pause
