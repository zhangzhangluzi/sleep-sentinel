#define MyAppName "SleepSentinel"
#define MyAppVersion "1.0.10"
#define MyAppPublisher "Zhang"
#define MyAppExeName "SleepSentinel.exe"

[Setup]
AppId={{2A9A948D-EEC7-4FB5-BD8A-8A5A63E55F13}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=SleepSentinel-Setup-win-x64
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=Assets\AppIcon.ico
CloseApplications=no
RestartApplications=no

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加选项:"

[Files]
Source: "bin\Release\net8.0-windows\win-x64\publish\SleepSentinel.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function IsSleepSentinelRunning(): Boolean;
var
  ResultCode: Integer;
begin
  Result := False;

  if not Exec(
    ExpandConstant('{cmd}'),
    '/C tasklist /FI "IMAGENAME eq {#MyAppExeName}" /NH | find /I "{#MyAppExeName}" >nul',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode) then
  begin
    Log('无法检查 SleepSentinel 进程状态。');
    Exit;
  end;

  Result := ResultCode = 0;
end;

function CloseSleepSentinelProcess(): Boolean;
var
  ResultCode: Integer;
  Attempt: Integer;
begin
  Result := True;

  if not IsSleepSentinelRunning() then
  begin
    Log('安装前未检测到正在运行的 SleepSentinel。');
    Exit;
  end;

  Log('安装前尝试关闭正在运行的 SleepSentinel。');
  if not Exec(
    ExpandConstant('{cmd}'),
    '/C taskkill /IM "{#MyAppExeName}" /T /F >nul 2>&1',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode) then
  begin
    Log('启动 taskkill 失败。');
    Result := False;
    Exit;
  end;

  for Attempt := 1 to 10 do
  begin
    if not IsSleepSentinelRunning() then
    begin
      Log('已成功关闭 SleepSentinel。');
      Exit;
    end;

    Sleep(500);
  end;

  Log('SleepSentinel 仍在运行，安装无法继续覆盖旧文件。');
  Result := False;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';

  if not CloseSleepSentinelProcess() then
  begin
    Result := '安装前无法关闭正在运行的 SleepSentinel，请先从托盘退出后重试。';
  end;
end;
