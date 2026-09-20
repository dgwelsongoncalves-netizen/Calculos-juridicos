[Setup]
AppName=Nash System
AppVersion=2.8.19
DefaultDirName={autopf}\Nash System
DefaultGroupName=Nash System
OutputDir=.\Instalador_Final
OutputBaseFilename=Instalador_Nash_System
SetupIconFile=dr_nash.ico
UninstallDisplayIcon={app}\nash.exe
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Area de Trabalho"; GroupDescription: "Atalhos adicionais:"

[Files]
Source: "dist\nash.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "Tabelas_Oficiais\*"; DestDir: "{app}\Tabelas_Oficiais"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dr_nash.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "template_nash.xlsx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Nash System"; Filename: "{app}\nash.exe"; IconFilename: "{app}\dr_nash.ico"
Name: "{autodesktop}\Nash System"; Filename: "{app}\nash.exe"; Tasks: desktopicon; IconFilename: "{app}\dr_nash.ico"