#define MyAppName "Local Word AI"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "LocalWordAI"
#define MyAppExeName "LocalWordAI.dll"
#define MyBuildDir "..\LocalWordAI\bin\Release"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\LocalWordAI
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=.\Output
OutputBaseFilename=LocalWordAI-Setup-v{#MyAppVersion}
SetupIconFile=
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
MinVersion=6.1sp1

[Languages]
Name: "vietnamese"; MessagesFile: "compiler:Default.isl"

[Files]
; Add-in DLL and dependencies
Source: "{#MyBuildDir}\LocalWordAI.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyBuildDir}\LocalWordAI.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MyBuildDir}\LocalWordAI.vsto"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MyBuildDir}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MyBuildDir}\Polly.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MyBuildDir}\DiffPlex.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#MyBuildDir}\FuzzySharp.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

; Skills mẫu
Source: "..\Skills\*.json"; DestDir: "{userappdata}\LocalWordAI\Skills"; Flags: ignoreversion

; Tài liệu hướng dẫn
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

; Đăng ký add-in với Word (HKCU - không cần admin)
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LocalWordAI"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LocalWordAI"; ValueType: string; ValueName: "Description"; ValueData: "Local AI Assistant for Microsoft Word - Offline"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LocalWordAI"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Local Word AI"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LocalWordAI"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\LocalWordAI"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\LocalWordAI.vsto|vstolocal"

[Code]
function IsDotNet48Installed(): Boolean;
var
  regKey: String;
  value: Cardinal;
begin
  regKey := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full';
  if RegQueryDWordValue(HKLM, regKey, 'Release', value) then
    Result := value >= 528040
  else
    Result := False;
end;

function IsVSTORuntimeInstalled(): Boolean;
var
  installed: String;
begin
  Result := RegQueryStringValue(HKLM,
    'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
    'Version', installed);
end;

procedure CleanupResiliencyAndProtect;
var
  key: String;
begin
  for key in [
    'Software\Microsoft\Office\16.0\Word\Resiliency',
    'Software\Microsoft\Office\15.0\Word\Resiliency'
  ] do
  begin
    RegDeleteKeyIncludingSubkeys(HKCU, key);
  end;

  RegWriteDWordValue(HKCU,
    'Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList',
    'LocalWordAI', 1);
end;

function InitializeSetup(): Boolean;
var
  vstoMissing: Boolean;
begin
  Result := True;

  if not IsDotNet48Installed() then
  begin
    MsgBox('Cần cài .NET Framework 4.8 trước khi cài Local Word AI.' + #13#10 +
           'Tải tại: https://dotnet.microsoft.com/download/dotnet-framework/net48',
           mbError, MB_OK);
    Result := False;
    Exit;
  end;

  vstoMissing := not IsVSTORuntimeInstalled();
  if vstoMissing then
  begin
    MsgBox('Cần cài VSTO Runtime trước khi cài Local Word AI.' + #13#10 +
           'Tìm file "vstor_redist.exe" trong thư mục Prerequisites và cài trước.' + #13#10 +
           'Hoặc chạy install.bat (tu dong cai VSTO Runtime neu co file).',
           mbInformation, MB_OK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    CleanupResiliencyAndProtect;
end;

[Run]
; Chạy regasm sau khi cài (nếu cần)
; Filename: "{dotnet40}\regasm.exe"; Parameters: "/codebase ""{app}\LocalWordAI.dll"""; Flags: runhidden

[UninstallRun]
; Xóa registry entries khi gỡ cài
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Microsoft\Office\Word\Addins\LocalWordAI"" /f"; Flags: runhidden

[Messages]
WelcomeLabel2=Local Word AI sẽ được cài đặt vào máy tính của bạn.%n%nAdd-in này chạy hoàn toàn offline và kết nối LM Studio local.%n%nHãy đảm bảo LM Studio đã được cài đặt và chạy trước khi sử dụng.
