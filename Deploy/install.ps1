# ============================================================
# Local Word AI - Script cai dat
# Chay bang: powershell -ExecutionPolicy Bypass -File install.ps1
# ============================================================

param(
    [switch]$SkipRuntimeCheck = $false,
    [switch]$ResetOnly = $false
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== Local Word AI - Cai dat / Khoi phuc ===" -ForegroundColor Cyan

# ── Mode: Reset Only ─────────────────────────────────────────
if ($ResetOnly) {
    Write-Host "`n[Reset] Chi khoi phuc registry..." -ForegroundColor Yellow
    $regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\LocalWordAI"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Type DWORD -Value 3
    Set-ItemProperty -Path $regPath -Name "Manifest" -Value ("file:///" + $scriptDir.Replace('\','/') + "/AddIn/LocalWordAI.vsto|vstolocal")
    Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Local Word AI"
    Set-ItemProperty -Path $regPath -Name "Description" -Value "Local AI Assistant for Word - Offline"

    Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item 'HKCU:\Software\Microsoft\Office\15.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue

    $doNotDisableKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList"
    if (-not (Test-Path $doNotDisableKey)) { New-Item -Path $doNotDisableKey -Force | Out-Null }
    Set-ItemProperty -Path $doNotDisableKey -Name "LocalWordAI" -Type DWORD -Value 1

    Write-Host "  LoadBehavior=3, da xoa toan bo Resiliency cache." -ForegroundColor Green
    Write-Host ""
    Write-Host "=== Khoi phuc hoan tat! Mo Word binh thuong. ===" -ForegroundColor Green
    Read-Host "Nhan Enter de thoat"
    exit 0
}

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
        Write-Host "  Dang cai VSTO Runtime (offline)..." -ForegroundColor Yellow
        Start-Process -FilePath $vstoInstaller -ArgumentList "/quiet /norestart" -Wait
        Write-Host "  VSTO Runtime da cai xong." -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "  LOI: Khong tim thay VSTO Runtime!" -ForegroundColor Red
        Write-Host "  May nay khong co internet. Can tai VSTO Runtime tu may khac:" -ForegroundColor Red
        Write-Host "    1. Tai: https://aka.ms/vstoruntime" -ForegroundColor Red
        Write-Host "    2. Luu file vstor_redist.exe vao cung thu muc voi install.bat" -ForegroundColor Red
        Write-Host "    3. Chay lai install.bat" -ForegroundColor Red
        Write-Host ""
        if (-not $SkipRuntimeCheck) {
            Read-Host "Nhan Enter de thoat"
            exit 1
        }
    }
}

# ── 2. Copy files ────────────────────────────────────────────
Write-Host "`n[2/5] Copy files..." -ForegroundColor Yellow
$installDir = "$env:LOCALAPPDATA\LocalWordAI"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Write-Host "  Thu muc: $installDir"

Get-ChildItem "$scriptDir\AddIn\*" -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item $_.FullName -Destination $installDir -Force
    Write-Host "  Copied: $($_.Name)"
}

# Xoa Zone.Identifier (Mark of the Web) khoi cac file vua copy
Get-ChildItem $installDir -ErrorAction SilentlyContinue | ForEach-Object {
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

# ── 5. Ngan Word tu dong disable add-in + Xoa toan bo Resiliency ──
Write-Host "`n[5/5] Ngan Word tu dong tat add-in..." -ForegroundColor Yellow

# DoNotDisableAddinList: danh sach add-in ma Office khong duoc tu dong disable
$doNotDisableKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList"
if (-not (Test-Path $doNotDisableKey)) { New-Item -Path $doNotDisableKey -Force | Out-Null }
Set-ItemProperty -Path $doNotDisableKey -Name "LocalWordAI" -Type DWORD -Value 1
Write-Host "  Da them vao DoNotDisableAddinList." -ForegroundColor Green

# Xoa toan bo cache Resiliency de Word khoi dau tu trang thai sach
$resilKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency",
    "HKCU:\Software\Microsoft\Office\15.0\Word\Resiliency"
)
foreach ($k in $resilKeys) {
    if (Test-Path $k) {
        Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Da xoa: $($k -replace 'HKCU:\\Software\\','')" -ForegroundColor Green
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
Write-Host "Luu y: Neu add-in van khong xuat hien, chay reset_word_ai.bat de fix LoadBehavior." -ForegroundColor Cyan
Write-Host ""
Read-Host "Nhan Enter de thoat"
