# ============================================================
# Local Word AI - Script chẩn đoán lỗi
# Double-click diagnose.bat để chạy
# Kết quả lưu tại Desktop\LocalWordAI_diagnose.txt
# ============================================================

$logPath = "$env:USERPROFILE\Desktop\LocalWordAI_diagnose.txt"
$lines   = [System.Collections.Generic.List[string]]::new()

function Log($msg) {
    $lines.Add($msg)
    Write-Host $msg
}

function Section($title) {
    Log ""
    Log ("=" * 60)
    Log "  $title"
    Log ("=" * 60)
}

Log "Local Word AI - Bao cao chan doan loi"
Log "Thoi gian: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log "May tinh:  $env:COMPUTERNAME"
Log "User:      $env:USERNAME"
Log "OS:        $([System.Environment]::OSVersion.VersionString)"

# ── 1. Registry ──────────────────────────────────────────────
Section "1. REGISTRY ADD-IN"
$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\LocalWordAI"
if (Test-Path $regPath) {
    $reg = Get-ItemProperty $regPath
    Log "LoadBehavior : $($reg.LoadBehavior)  (can = 3, neu = 2 thi Word da tu tat do loi)"
    Log "Manifest     : $($reg.Manifest)"
    Log "FriendlyName : $($reg.FriendlyName)"
} else {
    Log "LOI: Khong tim thay registry key!"
    Log "     Add-in chua duoc dang ky. Chay lai install.bat."
}

# ── 2. File manifest & DLL ───────────────────────────────────
Section "2. KIEM TRA FILE"
try {
    $manifest = $reg.Manifest -replace "\|vstolocal$","" -replace "^file:///",""
    $manifest = [Uri]::UnescapeDataString($manifest)
    $installDir = Split-Path $manifest

    $required = @(
        "LocalWordAI.vsto",
        "LocalWordAI.dll",
        "LocalWordAI.dll.manifest",
        "Newtonsoft.Json.dll",
        "Polly.dll",
        "DiffPlex.dll",
        "FuzzySharp.dll",
        "Microsoft.Office.Interop.Word.dll",
        "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
        "Microsoft.Vbe.Interop.dll",
        "Office.dll",
        "stdole.dll"
    )
    foreach ($f in $required) {
        $fp = Join-Path $installDir $f
        if (Test-Path $fp) {
            $size = (Get-Item $fp).Length
            Log "  [OK]    $f  ($size bytes)"
        } else {
            Log "  [THIEU] $f  <-- THIEU FILE NAY"
        }
    }
} catch {
    Log "Khong doc duoc manifest: $_"
}

# ── 3. Certificate ───────────────────────────────────────────
Section "3. CERTIFICATE (chu ky so)"
$thumbprint = "EBB4874F04D62BD6846A076DEA045E2A94E39244"

$tp = Get-ChildItem Cert:\CurrentUser\TrustedPublisher -ErrorAction SilentlyContinue |
      Where-Object { $_.Thumbprint -eq $thumbprint }
Log "TrustedPublisher : $(if ($tp) {'CO (OK)'} else {'CHUA CO -- Add-in se bi tu choi'})"

$rt = Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
      Where-Object { $_.Thumbprint -eq $thumbprint }
Log "Root (CA)        : $(if ($rt) {'CO (OK)'} else {'CHUA CO -- Co the gay loi trust'})"

# Thu cai lai cert neu co file .cer
$cerFile = Join-Path $installDir "LocalWordAI.cer" -ErrorAction SilentlyContinue
if (-not $tp -and (Test-Path $cerFile)) {
    Log "  -> Dang thu cai certificate tu $cerFile ..."
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerFile)
        $s1 = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher","CurrentUser")
        $s1.Open("ReadWrite"); $s1.Add($cert); $s1.Close()
        $s2 = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
        $s2.Open("ReadWrite"); $s2.Add($cert); $s2.Close()
        Log "  -> Certificate da duoc cai lai thanh cong!"
    } catch {
        Log "  -> Loi cai certificate: $_"
    }
}

# ── 4. VSTO Runtime ──────────────────────────────────────────
Section "4. VSTO RUNTIME"
$vstoKeys = @(
    "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R",
    "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4"
)
$vstoFound = $false
foreach ($k in $vstoKeys) {
    if (Test-Path $k) {
        $ver = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Version
        Log "  [OK] $k  version=$ver"
        $vstoFound = $true
    }
}
if (-not $vstoFound) {
    Log "  CHUA CAI VSTO Runtime -- Day la nguyen nhan chinh!"
    Log "  Tai: https://aka.ms/vstoruntime  (vstor_redist.exe ~38MB)"
}

# ── 5. .NET Framework ────────────────────────────────────────
Section "5. .NET FRAMEWORK"
$net48 = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
if ($net48) {
    $release = $net48.Release
    $version = $net48.Version
    $ok = $release -ge 528040
    Log ".NET 4.8: $version (Release=$release) -- $(if ($ok) {'OK'} else {'CAN CAP NHAT'})"
} else {
    Log ".NET 4.x: KHONG TIM THAY"
}

# ── 6. Office version ────────────────────────────────────────
Section "6. MICROSOFT OFFICE"
$officeKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot",
    "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot"
)
$officeFound = $false
foreach ($k in $officeKeys) {
    if (Test-Path $k) {
        $path = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Path
        Log "  Office path: $path"
        $officeFound = $true
    }
}
if (-not $officeFound) { Log "  Khong tim thay Office trong registry" }

# Word version
$wordExe = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
    "$env:ProgramFiles\Microsoft Office\Office16\WINWORD.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($wordExe) {
    $wv = (Get-Item $wordExe).VersionInfo.ProductVersion
    Log "  Word exe: $wordExe  (v$wv)"
} else {
    Log "  Khong tim thay WINWORD.EXE"
}

# ── 7. Word disabled items ───────────────────────────────────
Section "7. WORD DISABLED ADD-INS CACHE"
$disKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\15.0\Word\Resiliency\DisabledItems"
)
$disabledCount = 0
foreach ($dk in $disKeys) {
    if (Test-Path $dk) {
        $props = Get-ItemProperty $dk -ErrorAction SilentlyContinue
        $props.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
            $val = $_.Value
            if ($val -is [byte[]]) { $val = [System.Text.Encoding]::Unicode.GetString($val) }
            Log "  Disabled: $($_.Name) = $val"
            $disabledCount++
        }
    }
}
if ($disabledCount -eq 0) { Log "  (Khong co add-in nao bi disable trong cache)" }

# ── 8. Event Log ─────────────────────────────────────────────
Section "8. WINDOWS EVENT LOG (loi lien quan Word/VSTO)"
try {
    $events = Get-WinEvent -LogName Application -MaxEvents 500 -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.Message -match "LocalWordAI|VSTO|AddIn|Word" -or
             $_.ProviderName -match "VSTO|Word") -and
            $_.LevelDisplayName -match "Error|Warning"
        } |
        Select-Object -First 20

    if ($events) {
        foreach ($ev in $events) {
            Log "[$($ev.TimeCreated.ToString('HH:mm:ss'))] [$($ev.LevelDisplayName)] $($ev.ProviderName)"
            $msg = ($ev.Message -split "`n")[0..2] -join " | "
            Log "  $msg"
            Log ""
        }
    } else {
        Log "  (Khong tim thay loi lien quan trong Event Log)"
    }
} catch {
    Log "  Khong doc duoc Event Log: $_"
}

# ── 9. Tu dong reset LoadBehavior ────────────────────────────
Section "9. TU DONG SUA LOI"
if (Test-Path $regPath) {
    $lb = (Get-ItemProperty $regPath).LoadBehavior
    if ($lb -ne 3) {
        Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Type DWORD -Value 3
        Log "  Da reset LoadBehavior tu $lb -> 3"
    } else {
        Log "  LoadBehavior da la 3 (OK)"
    }
}

# Xoa disabled cache
foreach ($dk in $disKeys) {
    if (Test-Path $dk) {
        Remove-Item $dk -Recurse -Force -ErrorAction SilentlyContinue
        Log "  Da xoa Resiliency\DisabledItems cache"
    }
}

# ── Luu log ──────────────────────────────────────────────────
$lines | Set-Content -Path $logPath -Encoding UTF8
Write-Host ""
Write-Host "=== Log da luu tai: $logPath ===" -ForegroundColor Green
Write-Host "Gui file nay de kiem tra loi." -ForegroundColor Cyan
Write-Host ""

# Mo file log
Start-Process notepad $logPath

Read-Host "Nhan Enter de thoat"
