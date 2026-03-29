[Setup]
AppName=SSNnote
AppVersion=1.67
AppPublisher=justkevin2570
DefaultDirName={autopf}\SSNnote
DefaultGroupName=SSNnote
OutputDir=dist
OutputBaseFilename=SSNnote_Setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
CloseApplications=yes
AppId={{12345678-1234-1234-1234-123456789012}}
VersionInfoVersion=1.67.0.0
VersionInfoProductName=SSNnote
VersionInfoCompany=justkevin2570
VersionInfoProductVersion=1.67

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Files]
Source: "dist\SSNnote\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\SSNnote"; Filename: "{app}\SSNnote.exe"
Name: "{group}\{cm:UninstallProgram,SSNnote}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\SSNnote.exe"; Description: "{cm:LaunchProgram,SSNnote}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}"
