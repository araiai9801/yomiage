@echo off
chcp 65001 >nul
echo.
echo ========================================
echo   読み上げアプリ (yomiage) を起動します
echo ========================================
echo.
echo   Ctrl+Alt+R : 選択テキストを読み上げ
echo   Ctrl+Alt+E : カーソル位置から末尾まで読み上げ
echo   Esc        : 読み上げを停止
echo   トレイ右クリック : 停止 / 終了
echo.

:: pythonw.exe を自動検出
set PYTHONW=
for /f "usebackq tokens=*" %%i in (`where pythonw 2^>nul`) do (
    if not defined PYTHONW set PYTHONW=%%i
)

if not defined PYTHONW (
    echo [エラー] pythonw.exe が見つかりません。
    echo Python 3.10 以上をインストールし、PATH を通してください。
    echo https://www.python.org/downloads/windows/
    echo.
    pause
    exit /b 1
)

:: yomiage.py はこのバッチと同じフォルダにある想定
set SCRIPT=%~dp0yomiage.py

start "" "%PYTHONW%" "%SCRIPT%"
echo 起動しました。タスクバーのトレイアイコン「読」を確認してください。
echo.
pause
