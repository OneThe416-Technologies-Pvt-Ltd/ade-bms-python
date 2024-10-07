[Setup]
; Basic setup details
AppName=ADE BMS Software
AppVersion=1.0
DefaultDirName={commonpf}\ADE BMS Software
DefaultGroupName=ADE BMS Software
OutputDir=output
OutputBaseFilename=ADE BMS Software Installer
Compression=lzma
SolidCompression=yes

[Files]
; Add all the necessary files, including the .exe and any required assets
Source: "dist\main\ADE BMS.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\main\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Include driver files for both Windows 10 and 11
Source: "drivers\PeakOemDrv.exe"; DestDir: "{app}\drivers"; Flags: ignoreversion
Source: "drivers\driv_win_uport_v3.3_build_23100317_whql.exe"; DestDir: "{app}\drivers\win10"; Flags: ignoreversion
Source: "drivers\driv_win_uport_v4.1_build_23092610_whql.exe"; DestDir: "{app}\drivers\win11"; Flags: ignoreversion

[Icons]
; Create an icon for the application
Name: "{group}\ADE BMS Software"; Filename: "{app}\ADE BMS.exe"
; Create a desktop shortcut
Name: "{commondesktop}\ADE BMS Software"; Filename: "{app}\ADE BMS.exe"; Tasks: desktopicon

[Tasks]
; Define the tasks for optional components
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Run]
; Conditional installation of drivers based on Windows version
Filename: "{app}\drivers\PeakOemDrv.exe"; Parameters: ""; Description: "Install Common Drivers"; Flags: waituntilterminated runhidden
Filename: "{app}\drivers\win10\driv_win_uport_v3.3_build_23100317_whql.exe"; Parameters: ""; Description: "Install Windows 10 Drivers"; Flags: waituntilterminated runhidden; Check: IsWindows10
Filename: "{app}\drivers\win11\driv_win_uport_v4.1_build_23092610_whql.exe"; Parameters: ""; Description: "Install Windows 11 Drivers"; Flags: waituntilterminated runhidden; Check: IsWindows11
; Specify the file to run after installation
Filename: "{app}\ADE BMS.exe"; Description: "Launch ADE BMS Software"; Flags: nowait postinstall skipifsilent

[Code]
function IsWindows10: Boolean;
var
  Version: String;
begin
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'CurrentVersion', Version) then
    Result := (Version = '10.0')
  else
    Result := False;
end;

function IsWindows11: Boolean;
var
  BuildNumber: String;
begin
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows NT\CurrentVersion', 'CurrentBuild', BuildNumber) then
    Result := (StrToIntDef(BuildNumber, 0) >= 22000)
  else
    Result := False;
end;

function IsSupportedWindowsVersion: Boolean;
begin
  Result := IsWindows10 or IsWindows11;
end;

function InitializeSetup(): Boolean;
begin
  if not IsSupportedWindowsVersion then
  begin
    MsgBox('Warning: This application is only tested on Windows 10 and Windows 11. It may not work as expected on your operating system.', mbInformation, MB_OK);
  end;
  Result := True; // Continue with the installation
end;
