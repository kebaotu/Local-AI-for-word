@echo off
chcp 65001 >nul 2>&1
echo === Local Word AI - Khoi phuc LoadBehavior (fix add-in bi tat) ===
echo.

powershell -ExecutionPolicy Bypass -Command ^
    "Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue; ^
     Set-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\LocalDocAI' -Name LoadBehavior -Type DWORD -Value 3; ^
     Write-Host 'Hoan tat: LoadBehavior=3, da xoa cache Resiliency.' -ForegroundColor Green"

echo.
echo Hay mo Word binh thuong. Tab 'Local AI' se xuat hien.
pause
