; ============================================================
; Inno Setup script for Excel standardization Web App
;
; Prerequisites:
;   1. Build the exe first:  build_exe.bat
;   2. Install Inno Setup 6: https://jrsoftware.org/isinfo.php
;   3. Compile:              build_installer.bat
;                        or: iscc installer\Excelstandardization.iss
;
; Output: installer\Output\Excelstandardization_Setup_1.0.2.exe
; ============================================================

#define AppName      "Excel standardization"
#define AppVersion   "1.0.2"
#define AppPublisher "Excel standardization Team"
#define AppExeName   "Excelstandardization.exe"
#define DistDir      "..\dist\Excelstandardization"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename=Excelstandardization_Setup_{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Require admin rights so the app installs to Program Files
PrivilegesRequired=admin
; Minimum Windows version: Windows 10
MinVersion=10.0
; Architecture: x64 only
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";   Description: "Create a &desktop shortcut";    GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startmenuicon"; Description: "Create a &Start Menu shortcut"; GroupDescription: "Additional icons:"

[Files]
; Bundle the entire PyInstaller output folder
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
; Pre-create the writable runtime directories under %LOCALAPPDATA%\Excelstandardization
; so the app can write uploads/work/output files without needing admin rights at runtime.
Name: "{localappdata}\Excelstandardization";          Permissions: users-full
Name: "{localappdata}\Excelstandardization\uploads";  Permissions: users-full
Name: "{localappdata}\Excelstandardization\work";     Permissions: users-full
Name: "{localappdata}\Excelstandardization\output";   Permissions: users-full

[Icons]
; Start Menu shortcuts
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"; Tasks: startmenuicon
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}";      Tasks: startmenuicon
; Desktop shortcut
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Offer to launch the app after installation
Filename: "{app}\{#AppExeName}"; \
    Description: "Launch {#AppName} now"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove the app installation directory on uninstall
Type: filesandordirs; Name: "{app}"
; Leave user data in %LOCALAPPDATA%\Excelstandardization intact on uninstall
; (uploads and exports belong to the user — do not delete them silently)

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
