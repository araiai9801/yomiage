@echo off
chcp 65001 >nul
echo.
echo ========================================
echo   読み上げアプリ (yomiage) を起動します
echo ========================================
echo.
echo   Ctrl+Alt+R : 選択テキストを読み上げ
echo   Esc        : 読み上げを停止
echo   トレイ右クリック : 停止 / 終了
echo.
start "" "C:\Users\arai\AppData\Local\Programs\Python\Python313\pythonw.exe" "C:\Users\arai\OneDrive\ドキュメント\code\ClaudeCode\yomiage\yomiage.py"
echo 起動しました。タスクバーのトレイアイコン「読」を確認してください。
echo.
pause
