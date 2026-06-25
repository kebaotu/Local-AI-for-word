@echo off
echo === Local Word AI - Debug VSTO ===
echo.
echo LUU Y: Script nay dung de debug. Neu add-in chi bi LoadBehavior=2,
echo        hay dung reset_word_ai.bat (nhanh hon, khong can mo Word).
echo.

:: Dat bien moi truong VSTO log (TRUOC khi mo Word)
set VSTOHOST_LOGFILE=%USERPROFILE%\Desktop\VSTO_debug_log.txt
del "%VSTOHOST_LOGFILE%" 2>nul

:: Xoa tat ca Resiliency de Word thu lai (reset trang thai disable)
:: Word dat LoadBehavior=2 khi add-in gap loi. Xoa Resiliency + reset LoadBehavior=3
:: giup Word them mot lan load lai add-in.
powershell -ExecutionPolicy Bypass -Command ^
    "Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue; ^
     Remove-Item 'HKCU:\Software\Microsoft\Office\15.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue; ^
     Set-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\LocalWordAI' -Name LoadBehavior -Type DWORD -Value 3; ^
     Write-Host 'Da reset: LoadBehavior=3, da xoa toan bo Resiliency cache'"

echo.
echo DANG MO WORD voi debug log bat...
echo Vui long:
echo   1. Cho Word tai xong hoan toan (30 giay)
echo   2. Kiem tra xem tab "Local AI" co xuat hien khong
echo   3. DONG WORD (cua so nay se tu dong tiep tuc)
echo.

:: Mo Word va cho no dong (start /wait)
set WINWORD=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE
if not exist "%WINWORD%" (
    echo Tim kiem WINWORD...
    for /f "delims=" %%i in ('powershell -Command "Get-ChildItem 'C:\Program Files*' -Recurse -Filter WINWORD.EXE -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName"') do set WINWORD=%%i
)
start /wait "" "%WINWORD%"

echo.
echo Word da dong. Kiem tra ket qua...
echo.

:: Kiem tra LoadBehavior sau khi Word dong
powershell -ExecutionPolicy Bypass -Command ^
    "$lb = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\LocalWordAI').LoadBehavior; ^
     if ($lb -eq 3) { Write-Host 'LoadBehavior = 3 -> Add-in hoat dong OK!' -ForegroundColor Green } ^
     else { Write-Host \"LoadBehavior = $lb -> Add-in bi loi (Word tu tat)\" -ForegroundColor Red }"

echo.
echo === VSTO DEBUG LOG ===
if exist "%VSTOHOST_LOGFILE%" (
    echo Log file: %VSTOHOST_LOGFILE%
    echo.
    type "%VSTOHOST_LOGFILE%"
    echo.
    echo -> Mo log trong Notepad...
    start notepad "%VSTOHOST_LOGFILE%"
) else (
    echo Khong co VSTOHOST_LOGFILE (co the VSTO version nay dung registry log)
    echo Kiem tra: %USERPROFILE%\Desktop\VSTO_log.txt
    if exist "%USERPROFILE%\Desktop\VSTO_log.txt" (
        echo Tim thay VSTO_log.txt:
        type "%USERPROFILE%\Desktop\VSTO_log.txt"
        start notepad "%USERPROFILE%\Desktop\VSTO_log.txt"
    ) else (
        echo Khong co log nao.
        echo Hay gui LoadBehavior value sau khi dong Word.
    )
)

echo.
echo Ket qua kiem tra LoadBehavior se o tren.
echo Hay chup anh man hinh cua so nay va gui de kiem tra loi.
echo.
pause
