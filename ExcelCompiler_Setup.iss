; Script Inno Setup Intelligent pour Excel Compiler Pro
; Installateur unique avec détection automatique Windows 8.1/10/11
; Auteur: GOUNOU N'GOBI Chabi Zimé

[Setup]
; Informations de base
AppId={{A1B2C3D4-E5F6-7890-ABCD-123456789012}
AppName=Excel Compiler Pro
AppVersion=3.2
AppVerName=Excel Compiler Pro 3.2
AppPublisher=GOUNOU N'GOBI Chabi Zimé
AppPublisherURL=mailto:zimkada@gmail.com
AppSupportURL=mailto:zimkada@gmail.com
AppUpdatesURL=mailto:zimkada@gmail.com

; Répertoires d'installation
DefaultDirName={autopf}\Excel Compiler Pro
DefaultGroupName=Excel Compiler Pro
AllowNoIcons=yes
LicenseFile=
InfoBeforeFile=
InfoAfterFile=
OutputDir=installer
OutputBaseFilename=ExcelCompiler_Setup_Intelligent
; Informations détaillées
VersionInfoVersion=3.2.0.0
VersionInfoCompany=GOUNOU N'GOBI Chabi Zimé
VersionInfoDescription=Excel Compiler Pro - Outil professionnel de compilation Excel
VersionInfoCopyright=© 2026 GOUNOU N'GOBI Chabi Zimé
VersionInfoProductName=Excel Compiler Pro
VersionInfoProductVersion=3.2

; Icônes et images
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\ExcelCompiler.exe

; Compression
Compression=lzma2
SolidCompression=yes

; Privilèges et compatibilité
PrivilegesRequired=lowest
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

; Interface
WizardStyle=modern
DisableWelcomePage=no

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode
Name: "startmenu"; Description: "Ajouter au menu Démarrer"; GroupDescription: "{cm:AdditionalIcons}"
Name: "associate"; Description: "Associer les fichiers Excel (.xlsx, .xls, .csv)"; GroupDescription: "Associations de fichiers"; Flags: unchecked

[Files]
; Application modulaire ExcelCompiler v3.2 (PyQt6, Windows 10+)
; Build produit par : pyinstaller ExcelCompiler.spec  ->  dist\ExcelCompiler\
Source: "dist\ExcelCompiler\ExcelCompiler.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\ExcelCompiler\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Fichiers communs
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "folder.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "RELEASE_NOTES.md"; DestDir: "{app}"; Flags: ignoreversion

; Redistributable Microsoft Visual C++ 2015-2022 (requis par PyQt6 / Python 3.x)
Source: "redistributables\vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
; Menu Démarrer
Name: "{group}\Excel Compiler Pro"; Filename: "{app}\ExcelCompiler.exe"; IconFilename: "{app}\icon.ico"; Tasks: startmenu
Name: "{group}\Guide d'utilisation"; Filename: "{app}\README.md"; Tasks: startmenu
Name: "{group}\Notes de version"; Filename: "{app}\RELEASE_NOTES.md"; Tasks: startmenu
Name: "{group}\{cm:UninstallProgram,Excel Compiler Pro}"; Filename: "{uninstallexe}"; Tasks: startmenu
; Bureau
Name: "{autodesktop}\Excel Compiler Pro"; Filename: "{app}\ExcelCompiler.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon
; Barre de lancement rapide
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Excel Compiler Pro"; Filename: "{app}\ExcelCompiler.exe"; IconFilename: "{app}\icon.ico"; Tasks: quicklaunchicon

[Code]
// ExcelCompiler v3.2 cible Windows 10/11 (PyQt6). La version PyQt5/Win8 a été
// retirée (aucun build correspondant). On bloque proprement les OS antérieurs.
function IsWindows10Plus: Boolean;
var
  Version: TWindowsVersion;
begin
  GetWindowsVersionEx(Version);
  Result := Version.Major >= 10;  // Windows 10/11 = 10.0+
end;

// Vérification du redistributable Visual C++ 2015-2022.
function IsVCRedist2022Installed: Boolean;
begin
  Result := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64') or
            RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64');
end;

// Refuser l'installation sous Windows < 10.
function InitializeSetup: Boolean;
begin
  Result := IsWindows10Plus;
  if not Result then
    MsgBox('Excel Compiler Pro 3.2 nécessite Windows 10 ou 11.',
           mbCriticalError, MB_OK);
end;

[Run]
; Installation conditionnelle du redistributable Microsoft Visual C++ 2015-2022
Filename: "{tmp}\vc_redist.x64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installation des composants Microsoft Visual C++..."; Flags: waituntilterminated; Check: not IsVCRedist2022Installed
; Lancement de l'application
Filename: "{app}\ExcelCompiler.exe"; Description: "{cm:LaunchProgram,Excel Compiler Pro}"; Flags: nowait postinstall skipifsilent

[CustomMessages]
french.LaunchProgram=Lancer %1
english.LaunchProgram=Launch %1

[Registry]
; Associations de fichiers (optionnel)
Root: HKCR; Subkey: ".xlsx"; ValueType: string; ValueName: ""; ValueData: "ExcelCompilerFile"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; Subkey: ".xls"; ValueType: string; ValueName: ""; ValueData: "ExcelCompilerFile"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; Subkey: ".csv"; ValueType: string; ValueName: ""; ValueData: "ExcelCompilerFile"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; Subkey: "ExcelCompilerFile"; ValueType: string; ValueName: ""; ValueData: "Fichier Excel Compiler"; Flags: uninsdeletekey; Tasks: associate
Root: HKCR; Subkey: "ExcelCompilerFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\ExcelCompiler.exe,0"; Tasks: associate
Root: HKCR; Subkey: "ExcelCompilerFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\ExcelCompiler.exe"" ""%1"""; Tasks: associate

[UninstallDelete]
Type: filesandordirs; Name: "{app}"