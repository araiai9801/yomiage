# unregister_context_menu.ps1
# 右クリックメニューから「読み上げアプリ起動」を削除する

$regPath = "HKCU:\Software\Classes\Directory\Background\shell\yomiage"

if (Test-Path $regPath) {
    Remove-Item -Path $regPath -Recurse -Force
    Write-Output "右クリックメニューから削除しました"
} else {
    Write-Output "登録されていません（既に削除済み）"
}
