# ============================================================
# Local Word AI - Script cai dat
# Chay bang: powershell -ExecutionPolicy Bypass -File install.ps1
# ============================================================

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== Cai dat Local Word AI ===" -ForegroundColor Cyan

# ── 1. Kiem tra va cai VSTO Runtime neu can ──────────────────
Write-Host "`n[1/5] Kiem tra VSTO Runtime..." -ForegroundColor Yellow
$vstoKey  = "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
$vstoKey2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R"
$vstoFound = (Test-Path $vstoKey) -or (Test-Path $vstoKey2)

if ($vstoFound) {
    Write-Host "  VSTO Runtime da co san." -ForegroundColor Green
} else {
    $vstoInstaller = "$scriptDir\vstor_redist.exe"
    if (Test-Path $vstoInstaller) {
        Write-Host "  Dang cai VSTO Runtime..." -ForegroundColor Yellow
        Start-Process -FilePath $vstoInstaller -ArgumentList "/quiet /norestart" -Wait
        Write-Host "  VSTO Runtime da cai xong." -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "  CANH BAO: Thieu VSTO Runtime!" -ForegroundColor Red
        Write-Host "  Tai file vstor_redist.exe (64-bit, ~38MB) tu:" -ForegroundColor Red
        Write-Host "    https://aka.ms/vstoruntime" -ForegroundColor Red
        Write-Host "  Dat file vao cung thu muc voi install.bat, chay lai." -ForegroundColor Red
        Write-Host ""
        $choice = Read-Host "  Tiep tuc cai dat ma khong co VSTO Runtime? (y/n)"
        if ($choice -ne "y") { exit 1 }
    }
}

# ── 2. Copy files ────────────────────────────────────────────
Write-Host "`n[2/5] Copy files..." -ForegroundColor Yellow
$installDir = "$env:LOCALAPPDATA\LocalWordAI"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Write-Host "  Thu muc: $installDir"

Get-ChildItem "$scriptDir\AddIn\*" | ForEach-Object {
    Copy-Item $_.FullName -Destination $installDir -Force
    Write-Host "  Copied: $($_.Name)"
}

# Xoa Zone.Identifier (Mark of the Web) khoi cac file vua copy
Get-ChildItem $installDir | ForEach-Object {
    $zone = Get-Item $_.FullName -Stream "Zone.Identifier" -ErrorAction SilentlyContinue
    if ($zone) { Unblock-File -Path $_.FullName }
}

# ── 3. Cai certificate ───────────────────────────────────────
Write-Host "`n[3/5] Cai certificate..." -ForegroundColor Yellow
$cerPath = "$installDir\LocalWordAI.cer"

if (Test-Path $cerPath) {
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerPath)

        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher","CurrentUser")
        $store.Open("ReadWrite"); $store.Add($cert); $store.Close()

        $store2 = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
        $store2.Open("ReadWrite"); $store2.Add($cert); $store2.Close()

        Write-Host "  Certificate da cai." -ForegroundColor Green
    } catch {
        Write-Host "  LOI certificate: $_" -ForegroundColor Red
    }
} else {
    Write-Host "  Khong tim thay LocalWordAI.cer" -ForegroundColor Red
}

# ── 4. Dang ky add-in voi Word ───────────────────────────────
Write-Host "`n[4/5] Dang ky add-in voi Word..." -ForegroundColor Yellow
$vstoPath    = "$installDir\LocalWordAI.vsto"
$manifestUri = "file:///" + $vstoPath.Replace("\", "/") + "|vstolocal"
$regPath     = "HKCU:\Software\Microsoft\Office\Word\Addins\LocalWordAI"

if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "Manifest"     -Value $manifestUri
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Type DWORD -Value 3
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Local Word AI"
Set-ItemProperty -Path $regPath -Name "Description"  -Value "Local AI Assistant for Word - Offline"
Write-Host "  Dang ky: $manifestUri" -ForegroundColor Green

# ── 5. Ngan Word tu dong disable add-in ──────────────────────
Write-Host "`n[5/5] Ngan Word tu dong tat add-in..." -ForegroundColor Yellow

# DoNotDisableAddinList: danh sach add-in ma Office khong duoc tu dong disable
# Day la co che chinh thuc cua Microsoft - ngan Office set LoadBehavior = 2
$doNotDisableKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList"
if (-not (Test-Path $doNotDisableKey)) { New-Item -Path $doNotDisableKey -Force | Out-Null }
Set-ItemProperty -Path $doNotDisableKey -Name "LocalWordAI" -Type DWORD -Value 1
Write-Host "  LocalWordAI da duoc dua vao danh sach khong bi disable." -ForegroundColor Green

# Xoa Resiliency cache cu (phan khac, khong phai DoNotDisableAddinList)
$crashKey     = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\CrashingAddinList"
$disabledKey  = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems"
foreach ($k in @($crashKey, $disabledKey)) {
    if (Test-Path $k) {
        Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Da xoa cache cu: $(Split-Path $k -Leaf)" -ForegroundColor Green
    }
}

# ── Ket qua ──────────────────────────────────────────────────
Write-Host ""
Write-Host "=== Cai dat hoan tat! ===" -ForegroundColor Green
Write-Host ""
Write-Host "Buoc tiep theo:"
Write-Host "  1. Mo LM Studio, bat Local Server (port 1234)"
Write-Host "  2. Load mot model trong LM Studio"
Write-Host "  3. Mo Microsoft Word -> tab 'Local AI' se xuat hien"
Write-Host ""
Write-Host "Luu y: Mo Word binh thuong tu Start Menu hoac taskbar la duoc." -ForegroundColor Cyan
Write-Host "Neu chua thay tab: File > Options > Add-ins > COM Add-ins > Go" -ForegroundColor Yellow
Write-Host "                   Tick 'LocalWordAI', bam OK." -ForegroundColor Yellow
Write-Host ""
Read-Host "Nhan Enter de thoat"
