#define MyAppName "NetUpgrade"
#define MyAppPublisher "Innovate Nexus"
#ifndef AppVersion
  #define AppVersion "5.12.1"
#endif

[Setup]
AppId={{7A7E6D93-69B1-45D8-B566-8331B95D4AB8}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://innovatenexus.vip
AppSupportURL=https://github.com/charithads/NetUpgrade
AppUpdatesURL=https://github.com/charithads/NetUpgrade/releases
DefaultDirName={localappdata}\Programs\NetUpgrade
DefaultGroupName=NetUpgrade
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=NetUpgrade-Setup-v{#AppVersion}
SetupIconFile=app.ico
UninstallDisplayIcon={app}\NetUpgrade.exe
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=no
RestartApplications=no
Uninstallable=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
Source: "dist\NetUpgrade.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\NetUpgradeUpdater.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\NetUpgrade"; Filename: "{app}\NetUpgrade.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\NetUpgrade"; Filename: "{app}\NetUpgrade.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\NetUpgrade.exe"; Description: "Launch NetUpgrade"; WorkingDir: "{app}"; Flags: postinstall nowait skipifsilent

