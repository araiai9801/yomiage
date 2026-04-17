# register_context_menu.ps1
# デスクトップ右クリックメニューに「読み上げアプリ起動」を追加する
# 管理者権限で実行してください: 右クリック → 管理者として実行
#
# 削除したい場合: unregister_context_menu.ps1 を実行

# pythonw.exe を自動検出
$pythonw = (Get-Command pythonw -ErrorAction SilentlyContinue)?.Source
if (-not $pythonw) {
    Write-Error "pythonw.exe が見つかりません。Python 3.10 以上をインストールし PATH を通してください。"
    exit 1
}

# このスクリプトと同じフォルダの yomiage.py を使う
$script = Join-Path $PSScriptRoot "yomiage.py"
if (-not (Test-Path $script)) {
    Write-Error "yomiage.py が見つかりません: $script"
    exit 1
}

# デスクトップ背景の右クリックメニューに追加
$regPath = "HKCU:\Software\Classes\Directory\Background\shell\yomiage"
$cmdPath = "$regPath\command"

# 既存キーがあれば削除
if (Test-Path $regPath) {
    Remove-Item -Path $regPath -Recurse -Force
}

# メニュー項目を作成
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "(Default)" -Value "読み上げアプリを起動"
Set-ItemProperty -Path $regPath -Name "Icon"      -Value "$pythonw,0"

# コマンドを登録
New-Item -Path $cmdPath -Force | Out-Null
Set-ItemProperty -Path $cmdPath -Name "(Default)" -Value "`"$pythonw`" `"$script`""

Write-Output ""
Write-Output "========================================="
Write-Output "  右クリックメニューに登録しました"
Write-Output "========================================="
Write-Output ""
Write-Output "  デスクトップやフォルダの背景を右クリック →"
Write-Output "  「読み上げアプリを起動」 が表示されます"
Write-Output ""
Write-Output "  削除するには unregister_context_menu.ps1 を実行"
Write-Output ""
