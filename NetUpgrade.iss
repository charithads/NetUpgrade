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

[Registry]
Root: HKCU; Subkey: "Software\Innovate Nexus\NetUpgrade"; ValueType: string; ValueName: "ImagesPath"; ValueData: "{code:GetImagesPath}"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Innovate Nexus\NetUpgrade"; ValueType: string; ValueName: "BackupsPath"; ValueData: "{code:GetBackupsPath}"; Flags: uninsdeletevalue

[Run]
Filename: "{app}\NetUpgrade.exe"; Description: "Launch NetUpgrade"; WorkingDir: "{app}"; Flags: postinstall nowait skipifsilent


[Code]
var
  IOSStoragePage: TInputDirWizardPage;
  BackupsStoragePage: TInputDirWizardPage;

procedure InitializeWizard;
begin
  IOSStoragePage := CreateInputDirPage(wpWelcome,
    'Choose IOS image folder',
    'Select where IOS image files will be stored.',
    'NetUpgrade stores uploaded Cisco IOS-XE image files in this folder.',
    False, '');
  IOSStoragePage.Add('IOS image folder:');
  IOSStoragePage.Values[0] := ExpandConstant('{localappdata}\NetUpgrade\IOS Images');

  BackupsStoragePage := CreateInputDirPage(IOSStoragePage.ID,
    'Choose switch backup folder',
    'Select where switch configuration backups will be stored.',
    'NetUpgrade stores switch backups, reports, and upgrade run logs in this folder.',
    False, '');
  BackupsStoragePage.Add('Switch backup folder:');
  BackupsStoragePage.Values[0] := ExpandConstant('{localappdata}\NetUpgrade\Switch Configuration Backups');
end;

function GetImagesPath(Param: String): String;
begin
  Result := ExpandConstant(IOSStoragePage.Values[0]);
end;

function GetBackupsPath(Param: String): String;
begin
  Result := ExpandConstant(BackupsStoragePage.Values[0]);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if not ForceDirectories(IOSStoragePage.Values[0]) then
    Result := 'NetUpgrade could not create the selected IOS image folder.';
  if (Result = '') and not ForceDirectories(BackupsStoragePage.Values[0]) then
    Result := 'NetUpgrade could not create the selected switch backup folder.';
end;
