; -- Inno Setup Script --

[Setup]
; General settings
AppName=ADEBMS
AppVersion=1.0
DefaultDirName={commonpf}\ADEBMS
DefaultGroupName=ADEBMS
OutputDir=.
OutputBaseFilename=ADEBMSInstaller
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Application files
Source: "dist\main\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "files\ADEPCAN.sys"; DestDir: "{sys}\drivers";
Source: "files\ADEPCANBasic.dll"; DestDir: "{sys}";

[Run]
; Register the DLL (if needed)
Filename: "regsvr32"; Parameters: "/s {sys}\ADEPCANBasic.dll"; Flags: runhidden
; Install the driver using pnputil
Filename: "{sys}\pnputil.exe"; Parameters: "/add-driver {sys}\drivers\ADEPCAN.sys /install"; Flags: runhidden waituntilterminated

[Icons]
; Create shortcuts
Name: "{group}\ADEBMS"; Filename: "{app}\main.exe"
Name: "{commondesktop}\ADEBMS"; Filename: "{app}\main.exe"

[Code]
function InitializeSetup(): Boolean;
begin
  // Check for administrator privileges
  if not IsAdmin then
  begin
    MsgBox('This installer requires administrator privileges. Please run as administrator.', mbError, MB_OK);
    Result := False;
  end
  else
    Result := True;
end;
