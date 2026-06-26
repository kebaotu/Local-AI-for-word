# Build script for Local Word AI
# Yêu cầu: Visual Studio 2019/2022 hoặc Build Tools, NuGet CLI

param(
    [string]$Configuration = "Release",
    [switch]$Package = $false
)

$ErrorActionPreference = "Stop"
$ProjectDir = "$PSScriptRoot\LocalDocAI"
$SlnFile = "$PSScriptRoot\LocalDocAI.sln"

Write-Host "=== Local Word AI Build ===" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration"

# 1. Restore NuGet packages
Write-Host "`n[1] Restoring NuGet packages..." -ForegroundColor Yellow
$nuget = Get-Command nuget -ErrorAction SilentlyContinue
if ($nuget) {
    & nuget restore $SlnFile
} else {
    Write-Host "nuget.exe not found - trying dotnet restore" -ForegroundColor Gray
    & dotnet restore $SlnFile 2>$null
}

# 2. Build
Write-Host "`n[2] Building..." -ForegroundColor Yellow

# Find MSBuild
$msbuild = $null
$candidates = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)
foreach ($c in $candidates) {
    if (Test-Path $c) { $msbuild = $c; break }
}

if (-not $msbuild) {
    Write-Host "MSBuild not found! Install Visual Studio 2019/2022 or Build Tools." -ForegroundColor Red
    exit 1
}

Write-Host "Using MSBuild: $msbuild"
& $msbuild $SlnFile /p:Configuration=$Configuration /p:Platform="Any CPU" /t:Build /v:minimal

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build FAILED!" -ForegroundColor Red
    exit 1
}

Write-Host "`nBuild SUCCESS!" -ForegroundColor Green
Write-Host "Output: $ProjectDir\bin\$Configuration\"

# 3. Copy skills to AppData (optional dev step)
Write-Host "`n[3] Copying sample skills..." -ForegroundColor Yellow
$skillsDest = "$env:APPDATA\LocalDocAI\Skills"
if (-not (Test-Path $skillsDest)) { New-Item -ItemType Directory -Force $skillsDest | Out-Null }
Copy-Item "$PSScriptRoot\Skills\*.json" $skillsDest -Force
Write-Host "Skills copied to: $skillsDest"

# 4. Package with Inno Setup (optional)
if ($Package) {
    Write-Host "`n[4] Packaging with Inno Setup..." -ForegroundColor Yellow
    $iscc = Get-Command iscc -ErrorAction SilentlyContinue
    if (-not $iscc) {
        $iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    }
    if (Test-Path $iscc) {
        & $iscc "$PSScriptRoot\Installer\LocalDocAI.iss"
        Write-Host "Package created in: $PSScriptRoot\Installer\Output\"
    } else {
        Write-Host "Inno Setup not found, skipping packaging." -ForegroundColor Gray
    }
}

Write-Host "`n=== Done ===" -ForegroundColor Cyan
