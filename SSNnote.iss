; 서서니 노트 설치 프로그램 (Inno Setup 6)
;
; 내 사용자 폴더(%LOCALAPPDATA%\Programs\SSNnote)에 설치한다 → 업데이트할 때 관리자 확인(UAC) 창이 뜨지 않음.
; 예전 판(v2.82까지)은 Program Files에 관리자 권한으로 설치했으므로, 처음 한 번 그 판을 지운다(이때만 UAC 한 번).
;
; 자동 업데이트(updater.py)는 /VERYSILENT ... 로 화면 없이 실행한다.
;   /NORELAUNCH  설치 뒤 앱을 다시 켜지 않음 (사용자가 앱을 끄면서 설치한 경우)
; 앱이 아직 끝나는 중이면(단일 실행 뮤텍스) 최대 15초 기다린다.

[Setup]
AppName=SSNnote
AppVersion=3.02
AppPublisher=justkevin2570
AppId={{12345678-1234-1234-1234-123456789012}}
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\SSNnote
DefaultGroupName=SSNnote
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableFinishedPage=yes
UsePreviousAppDir=no
OutputDir=dist
OutputBaseFilename=SSNnote_Setup
Compression=lzma
SolidCompression=yes
CloseApplications=force
RestartApplications=no
UninstallDisplayIcon={app}\SSNnote.exe
VersionInfoVersion=3.02.0.0
VersionInfoProductName=SSNnote
VersionInfoCompany=justkevin2570
VersionInfoProductVersion=3.02

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Files]
Source: "dist\SSNnote\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\SSNnote"; Filename: "{app}\SSNnote.exe"
Name: "{group}\{cm:UninstallProgram,SSNnote}"; Filename: "{uninstallexe}"

[Run]
; 설치가 끝나면 앱을 켠다 (화면 있는 설치·자동 업데이트 모두). /NORELAUNCH면 켜지 않음.
Filename: "{app}\SSNnote.exe"; Flags: nowait postinstall; Check: ShouldRelaunch

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueName: "서서니 메모"; Flags: uninsdeletevalue

[InstallDelete]
Type: files; Name: "{app}\_internal"
Type: files; Name: "{app}\null"
Type: files; Name: "{app}\SSNnote.exe"
Type: files; Name: "{app}\SSNnote_MCP.exe"

[UninstallDelete]
Type: dirifempty; Name: "{app}"

[Code]
const
  APP_MUTEX = 'Global\SSNnote_SingleInstance_9f3a2c';
  RUN_KEY = 'Software\Microsoft\Windows\CurrentVersion\Run';
  RUN_NAME = '서서니 메모';
  OLD_UNINSTALL_KEY = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{12345678-1234-1234-1234-123456789012}}_is1';  // AppId={{...}} 는 앞의 두 겹만 한 겹이 되어 끝은 두 겹 그대로

var
  HadAutostart: Boolean;

function ShouldRelaunch(): Boolean;
begin
  Result := Pos('/NORELAUNCH', Uppercase(GetCmdTail)) = 0;
end;

function InitializeSetup(): Boolean;
var
  i: Integer;
  v: String;
begin
  { 자동 업데이트: 앱이 완전히 끝날 때까지 잠깐 기다림 (최대 15초, 그 뒤엔 설치 프로그램이 닫음) }
  i := 0;
  while CheckForMutexes(APP_MUTEX) and (i < 60) do
  begin
    Sleep(250);
    i := i + 1;
  end;
  { 예전 판을 지우면 자동 시작 등록도 지워지므로, 켜져 있었는지 기억해 둔다 }
  HadAutostart := RegQueryStringValue(HKCU, RUN_KEY, RUN_NAME, v);
  Result := True;
end;

{ 예전 판(Program Files, 관리자 설치)의 제거 프로그램 경로 }
function OldAdminUninstaller(var Path: String): Boolean;
begin
  Result := RegQueryStringValue(HKLM32, OLD_UNINSTALL_KEY, 'UninstallString', Path);
  if (not Result) and IsWin64 then
    Result := RegQueryStringValue(HKLM64, OLD_UNINSTALL_KEY, 'UninstallString', Path);
  if Result then
    Path := RemoveQuotes(Path);
  Result := Result and FileExists(Path);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  p: String;
  rc: Integer;
begin
  if CurStep = ssInstall then
  begin
    { 예전 관리자 설치판을 조용히 지운다 — 관리자 확인(UAC)이 이때 한 번 뜬다.
      거절해도 새 판 설치는 계속하고, 다음 업데이트 때 다시 시도한다. 메모 데이터(%APPDATA%\SSNnote)는 지우지 않음. }
    if OldAdminUninstaller(p) then
      ShellExec('runas', p, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, rc);
  end;
  if CurStep = ssPostInstall then
  begin
    { 자동 시작이 켜져 있었으면 새 위치로 다시 등록 }
    if HadAutostart then
      RegWriteStringValue(HKCU, RUN_KEY, RUN_NAME, '"' + ExpandConstant('{app}\SSNnote.exe') + '"');
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    DelTree(ExpandConstant('{app}'), True, True, True);
end;
