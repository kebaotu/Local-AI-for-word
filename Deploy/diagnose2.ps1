# ============================================================
# Local Word AI - Chan doan sau (Phase 2)
# Chay script nay, sau do MO WORD, sau do DONG WORD
# Ket qua luu tai Desktop\LocalWordAI_diagnose2.txt
# ============================================================

$logPath = "$env:USERPROFILE\Desktop\LocalWordAI_diagnose2.txt"
$installDir = "$env:LOCALAPPDATA\LocalWordAI"
$lines = [System.Collections.Generic.List[string]]::new()

function Log($msg) { $lines.Add($msg); Write-Host $msg }
function Section($t) { Log ""; Log ("=" * 60); Log "  $t"; Log ("=" * 60) }

Log "Local Word AI - Chan doan sau (Phase 2)"
Log "Thoi gian: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# ── 1. Unblock files (Mark of the Web) ──────────────────────
Section "1. UNBLOCK FILES (Mark of the Web)"
try {
    $files = Get-ChildItem $installDir -ErrorAction Stop
    foreach ($f in $files) {
        $blocked = Get-Item $f.FullName -Stream "Zone.Identifier" -ErrorAction SilentlyContinue
        if ($blocked) {
            Unblock-File -Path $f.FullName
            Log "  Unblocked: $($f.Name)"
        } else {
            Log "  OK (not blocked): $($f.Name)"
        }
    }
} catch {
    Log "  Loi: $_"
}

# ── 2. Trust Center settings ─────────────────────────────────
Section "2. WORD TRUST CENTER SETTINGS"
$tcKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Word\Security",
    "HKCU:\Software\Policies\Microsoft\Office\16.0\Word\Security"
)
foreach ($k in $tcKeys) {
    if (Test-Path $k) {
        Log "  Key: $k"
        $props = Get-ItemProperty $k -ErrorAction SilentlyContinue
        @("VBAWarnings","ExtensionHardening","LoadCOMAddInsInNEW","DisableAllAddins",
          "RequireAddinSig","AllowAllTrustedLocations") | ForEach-Object {
            $v = $props.$_
            if ($null -ne $v) { Log "    $_ = $v" }
        }
    }
}

# Kiem tra neu co policy tat add-in
$disableKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Security"
$disable = (Get-ItemProperty $disableKey -ErrorAction SilentlyContinue).DisableAllAddins
if ($disable -eq 1) {
    Log "  CANH BAO: DisableAllAddins = 1 -> Tat het moi add-in!"
    Log "  -> Dang sua: DisableAllAddins = 0"
    Set-ItemProperty $disableKey -Name "DisableAllAddins" -Value 0 -Type DWORD
}

# ── 3. Kiem tra dependency DLL ───────────────────────────────
Section "3. KIEM TRA DEPENDENCY (thu load DLL)"
try {
    Add-Type -Path "$installDir\Newtonsoft.Json.dll" -ErrorAction Stop
    Log "  Newtonsoft.Json.dll: OK"
} catch { Log "  Newtonsoft.Json.dll: LOI - $_" }

try {
    Add-Type -Path "$installDir\Polly.dll" -ErrorAction Stop
    Log "  Polly.dll: OK"
} catch { Log "  Polly.dll: LOI - $_" }

try {
    Add-Type -Path "$installDir\DiffPlex.dll" -ErrorAction Stop
    Log "  DiffPlex.dll: OK"
} catch { Log "  DiffPlex.dll: LOI - $_" }

try {
    Add-Type -Path "$installDir\FuzzySharp.dll" -ErrorAction Stop
    Log "  FuzzySharp.dll: OK"
} catch { Log "  FuzzySharp.dll: LOI - $_" }

# ── 4. Kiem tra VSTO 64-bit vs 32-bit ───────────────────────
Section "4. VSTO RUNTIME 64-BIT vs 32-BIT"
$vsto64 = Test-Path "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
$vsto32 = Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R"
Log "  VSTO 64-bit (HKLM\SOFTWARE\Microsoft\...): $(if ($vsto64) {'CO (OK)'} else {'KHONG CO'})"
Log "  VSTO 32-bit (HKLM\WOW6432Node\...):        $(if ($vsto32) {'CO'} else {'KHONG CO'})"

# Word 365 64-bit can VSTO 64-bit
$wordExe = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
    "$env:ProgramFiles\Microsoft Office\Office16\WINWORD.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($wordExe) {
    $is64 = [System.Reflection.AssemblyName]::GetAssemblyName($wordExe) -eq $null
    $wordBits = if ($wordExe -match "Program Files \(x86\)") { "32-bit" } else { "64-bit (can VSTO 64-bit)" }
    Log "  Word: $wordExe"
    Log "  Word architecture: $wordBits"
    if ($wordBits -match "64-bit" -and -not $vsto64) {
        Log ""
        Log "  *** PHAT HIEN VAN DE ***"
        Log "  Word 64-bit nhung chi co VSTO 32-bit (WOW6432Node)"
        Log "  Day co the la nguyen nhan chinh khien add-in khong load!"
        Log ""
        Log "  Kiem tra thu muc VSTO Runtime 64-bit cua Office C2R..."
        $c2rVsto = @(
            "$env:ProgramFiles\Common Files\Microsoft Shared\VSTO\10.0",
            "$env:ProgramFiles\Microsoft Office\root\vfs\ProgramFilesX86\Common Files\Microsoft Shared\VSTO\10.0"
        )
        foreach ($p in $c2rVsto) {
            if (Test-Path $p) { Log "  Tim thay VSTO C2R tai: $p" }
        }
    }
}

# ── 4b. Bat VSTO diagnostic log ──────────────────────────────
Section "4b. BAT VSTO DIAGNOSTIC LOG"
$vstoLogPath = "$env:USERPROFILE\Desktop\VSTO_log.txt"
$vstoLogKey  = "HKCU:\Software\Microsoft\VSTO\Log"
if (-not (Test-Path $vstoLogKey)) { New-Item $vstoLogKey -Force | Out-Null }
Set-ItemProperty $vstoLogKey -Name "EnableLog"  -Value 1    -Type DWORD
Set-ItemProperty $vstoLogKey -Name "LogFilePath" -Value $vstoLogPath -Type String
Log "  VSTO log bat tai: $vstoLogPath"
Log "  -> Hay mo Word, dung 30 giay, dong Word, roi chay diagnose2.bat lan 2"

# ── 5. Reset LoadBehavior ────────────────────────────────────
Section "5. RESET LOADBEHAVIOR"
$regPath = "HKCU:\Software\Microsoft\Office\Word\Addins\LocalWordAI"
Set-ItemProperty $regPath -Name "LoadBehavior" -Type DWORD -Value 3
Log "  LoadBehavior = 3 (da reset)"

# Xoa tat ca Resiliency cache
$resilKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency",
    "HKCU:\Software\Microsoft\Office\15.0\Word\Resiliency"
)
foreach ($rk in $resilKeys) {
    if (Test-Path $rk) {
        Remove-Item $rk -Recurse -Force -ErrorAction SilentlyContinue
        Log "  Da xoa Resiliency cache: $rk"
    }
}

# ── 6. Doc VSTO log neu da co ────────────────────────────────
Section "6. VSTO LOG (neu da mo Word truoc khi chay script nay)"
if (Test-Path $vstoLogPath) {
    Log "  VSTO log tim thay:"
    Get-Content $vstoLogPath | ForEach-Object { Log "  $_" }
} else {
    Log "  Chua co VSTO log."
    Log "  -> Mo Word, dung 30 giay, dong Word"
    Log "  -> Sau do chay lai diagnose2.bat de xem log"
}

# ── 7. Event Log (sau khi mo Word) ──────────────────────────
Section "7. EVENT LOG (2 phut gan nhat)"
try {
    $cutoff = (Get-Date).AddMinutes(-120)
    $events = Get-WinEvent -LogName Application -MaxEvents 1000 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.TimeCreated -gt $cutoff -and
            ($_.Message -match "LocalWordAI|VSTO|AddIn|vstolocal|addin" -or
             $_.ProviderName -match "VSTO|Word|Office") -and
            $_.LevelDisplayName -match "Error|Warning|Critical"
        } | Select-Object -First 30

    if ($events) {
        foreach ($ev in $events) {
            Log "[$($ev.TimeCreated.ToString('HH:mm:ss'))] [$($ev.LevelDisplayName)] $($ev.ProviderName)"
            ($ev.Message -split "`n") | Select-Object -First 5 | ForEach-Object { Log "  $_" }
            Log ""
        }
    } else {
        Log "  Khong co event log loi nao trong 2 gio gan day"
    }
} catch { Log "  Loi doc event log: $_" }

# ── Luu log ──────────────────────────────────────────────────
$lines | Set-Content -Path $logPath -Encoding UTF8
Write-Host ""
Write-Host "Buoc tiep theo:" -ForegroundColor Yellow
Write-Host "  1. Script nay da bat VSTO diagnostic log va reset cai dat"
Write-Host "  2. Mo Word (cho 30 giay), dong lai"
Write-Host "  3. Chay lai diagnose2.bat lan 2"
Write-Host "  4. Gui ca 2 file: $logPath"
Write-Host "             va: $vstoLogPath (neu co)"
Write-Host ""
Write-Host "Log da luu: $logPath" -ForegroundColor Green
Start-Process notepad $logPath
Read-Host "Nhan Enter de thoat"
