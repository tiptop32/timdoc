#define MyAppName "Timdoc"
#define MyAppVersion "0.1.0"
#define MyAppExeName "Timdoc.exe"

[Setup]
AppId={{39BEF5F1-4EDC-4E53-9E03-763A7B08B1C3}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\Timdoc
DefaultGroupName=Timdoc
OutputDir=..\dist
OutputBaseFilename=Timdoc-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern

[Files]
Source: "..\dist\Timdoc\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Timdoc"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Timdoc"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Ярлыки:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить Timdoc"; Flags: nowait postinstall skipifsilent
