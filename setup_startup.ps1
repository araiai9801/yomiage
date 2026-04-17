# setup_startup.ps1
# yomiage をタスクスケジューラに登録（ログオン時に30秒遅延で起動）
# 旧スタートアップショートカットがあれば削除

$oldLnk = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\yomiage.lnk"
if (Test-Path $oldLnk) {
    Remove-Item $oldLnk -Force
    Write-Output "旧ショートカットを削除しました"
}

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

$taskName = "yomiage"

# 既存タスクがあれば削除
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action   = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`""
$trigger  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$trigger.Delay = "PT30S"  # 30秒遅延
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 0)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "選択テキスト読み上げ (yomiage)" -Force

Write-Output ""
Write-Output "========================================="
Write-Output "  タスクスケジューラに登録しました"
Write-Output "========================================="
Write-Output "  タスク名   : $taskName"
Write-Output "  pythonw    : $pythonw"
Write-Output "  スクリプト : $script"
Write-Output "  起動タイミング: ログオン 30 秒後"
Write-Output ""
Write-Output "  解除するには: タスクスケジューラで 'yomiage' タスクを削除"
Write-Output ""
