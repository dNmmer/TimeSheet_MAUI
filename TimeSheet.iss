#define MyAppName "TimeSheet"
#define MyAppVersion "1.0.0"
#define MyAppExeName "TimeSheet.exe"
#define MyAppPublisher ""
#define PublishDir "C:\\Project\\TimeSheet_app\\TimeSheet_MAUI\\publish\\win-x64"

[Setup]
AppId={{B3C7C50E-6C5B-4E3E-9A1A-FAE9C7A5C0F1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=no
OutputDir=installer
OutputBaseFilename=TimeSheet_{#MyAppVersion}_Setup
SetupIconFile=Platforms\Windows\appicon.ico
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык"; GroupDescription: "Дополнительные значки"; Flags: unchecked

[Files]
Source: "{#PublishDir}\\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"
Name: "{autodesktop}\\TimeSheet"; Filename: "{app}\\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent
