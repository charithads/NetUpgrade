#define MyAppName "NetUpgrade"
#define MyAppPublisher "Innovate Nexus"
#ifndef AppVersion
  #define AppVersion "5.16.0"
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
Root: HKCU; Subkey: "Software\Innovate Nexus\NetUpgrade"; ValueType: string; ValueName: "StoragePath"; ValueData: "{code:GetStoragePath}"; Flags: uninsdeletevalue

[Run]
Filename: "{app}\NetUpgrade.exe"; Description: "Launch NetUpgrade"; WorkingDir: "{app}"; Flags: postinstall nowait skipifsilent


[Code]
var
  StoragePage: TInputDirWizardPage;

procedure InitializeWizard;
begin
  StoragePage := CreateInputDirPage(wpWelcome,
    'Choose storage location',
    'Select one location for NetUpgrade files.',
    'Images, switch backups, reports, and upgrade logs will be kept under a NetUpgrade folder here. Use Make New Folder in Browse to create that folder if needed.',
    False, 'NetUpgrade');
  StoragePage.Add('Storage location:');
  StoragePage.Values[0] := ExpandConstant('{localappdata}');
end;

function GetStoragePath(Param: String): String;
var
  SelectedPath: String;
begin
  SelectedPath := ExpandConstant(Trim(StoragePage.Values[0]));
  while (Length(SelectedPath) > 3) and (Copy(SelectedPath, Length(SelectedPath), 1) = '\') do
    Delete(SelectedPath, Length(SelectedPath), 1);
  if CompareText(ExtractFileName(SelectedPath), 'NetUpgrade') = 0 then
    Result := SelectedPath
  else
    Result := AddBackslash(SelectedPath) + 'NetUpgrade';
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if Trim(StoragePage.Values[0]) = '' then
    Result := 'Select a storage location before continuing.'
  else if not ForceDirectories(GetStoragePath('')) then
    Result := 'NetUpgrade could not create the selected storage folder.';
end;
