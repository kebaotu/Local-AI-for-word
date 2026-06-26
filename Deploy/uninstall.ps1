# ============================================================
# Local Word AI - Script gỡ cài đặt
# ============================================================

Write-Host "=== Go cai dat Local Word AI ===" -ForegroundColor Cyan

# Xóa registry
$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\LocalDocAI"
if (Test-Path $regPath) {
    Remove-Item -Path $regPath -Force
    Write-Host "  Registry key removed" -ForegroundColor Green
}

# Xóa files
$installDir = "$env:LOCALAPPDATA\LocalDocAI"
if (Test-Path $installDir) {
    Remove-Item -Path $installDir -Recurse -Force
    Write-Host "  Files removed: $installDir" -ForegroundColor Green
}

Write-Host "Go cai dat hoan tat." -ForegroundColor Green
Read-Host "Nhan Enter de thoat"
