;
; Script to install Conform Universal
;

;; Pre-define ISPP variables
#define FileHandle
#define FileLine
#define MyInformationVersion

; Read the informational SEMVER version string from the file created by the build process
#define FileHandle = FileOpen("..\publish\InformationVersion.txt"); 
#define FileLine = FileRead(FileHandle)
#pragma message "Informational version number: " + FileLine

; Save the SEMVER version for use in the installer filename
#define MyInformationVersion FileLine

; Close the SEMVER version file
#if FileHandle
  #expr FileClose(FileHandle)
#endif

#define MyAppName "ASCOM Conform Universal"
#define MyAppPublisher "ASCOM Initiative (Peter Simpson)"
#define MyAppPublisherURL "https://ascom-standards.org"
#define MyAppSupportURL "URL=https://ascomtalk.groups.io/g/Developer/topics"
#define MyAppUpdatesURL "https://github.com/ASCOMInitiative/ConformU/releases"
#define MyAppExeName "ConformU.exe"
#define MyAppAuthor "Peter Simpson"
#define MyAppCopyright "Copyright © 2022 " + MyAppAuthor
#define MyAppVersion GetVersionNumbersString("..\publish\ConformU64\conformu.exe")  ; Create version number variable

[Setup]
AppId={{454081A3-59A9-4EC6-984C-11CA605D3196}
AppCopyright={#MyAppCopyright}
AppName={#MyAppName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppPublisherURL}
AppSupportURL={#MyAppSupportURL}
AppUpdatesURL={#MyAppUpdatesURL}
AppVerName={#MyAppName}
AppVersion={#MyAppVersion}
ArchitecturesInstallIn64BitMode=x64
Compression=lzma2/max
DefaultDirName={autopf}\ASCOM\ConformU
DefaultGroupName=ASCOMConformUniversal
DisableDirPage=yes
DisableProgramGroupPage=yes
MinVersion=6.1SP1
OutputBaseFilename=ConformU({#MyInformationVersion})Setup
OutputDir=.\Builds
PrivilegesRequired=admin
SetupIconFile=ASCOM.ico
SetupLogging=true
ShowLanguageDialog=auto
SignToolRunMinimized=yes
SignTool = SignConformU
SolidCompression=no
UninstallDisplayName=
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoCompany=ASCOM Initiative
VersionInfoCopyright={#MyAppAuthor}
VersionInfoDescription= {#MyAppName}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion= {#MyAppVersion}
VersionInfoVersion={#MyAppVersion}
WizardImageFile=NewWizardImage.bmp
WizardSmallImageFile=ASCOMLogo.bmp
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "armenian"; MessagesFile: "compiler:Languages\Armenian.isl"
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "bulgarian"; MessagesFile: "compiler:Languages\Bulgarian.isl"
Name: "catalan"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "corsican"; MessagesFile: "compiler:Languages\Corsican.isl"
Name: "czech"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "danish"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "finnish"; MessagesFile: "compiler:Languages\Finnish.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"
Name: "icelandic"; MessagesFile: "compiler:Languages\Icelandic.isl"
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
Name: "norwegian"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "slovak"; MessagesFile: "compiler:Languages\Slovak.isl"
Name: "slovenian"; MessagesFile: "compiler:Languages\Slovenian.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Files]
; 64bit OS - Install the 64bit app
Source: "..\publish\ConformU64\*.exe"; DestDir: "{app}"; Flags: ignoreversion signonce; Check: Is64BitInstallMode
Source: "..\publish\ConformU64\*.dll"; DestDir: "{app}"; Flags: ignoreversion signonce; Check: Is64BitInstallMode
Source: "..\publish\ConformU64\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes:"*.exe,*.dll"; Check: Is64BitInstallMode
Source: "..\publish\ConformU64\wwwroot\*"; DestDir: "{app}\wwwroot"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: Is64BitInstallMode

; 64bit OS - Install the 32bit app
Source: "..\publish\ConformU86\*.exe"; DestDir: "{app}\32bit"; Flags: ignoreversion signonce; Check: Is64BitInstallMode
Source: "..\publish\ConformU86\*.dll"; DestDir: "{app}\32bit"; Flags: ignoreversion signonce; Check: Is64BitInstallMode
Source: "..\publish\ConformU86\*"; DestDir: "{app}\32bit"; Flags: ignoreversion; Excludes:"*.exe,*.dll"; Check: Is64BitInstallMode
Source: "..\publish\ConformU86\wwwroot\*"; DestDir: "{app}\32bit\wwwroot"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: Is64BitInstallMode

; 32bit OS - Install the 32bit app
Source: "..\publish\ConformU86\*.exe"; DestDir: "{app}"; Flags: ignoreversion signonce; Check: not Is64BitInstallMode
Source: "..\publish\ConformU86\*.dll"; DestDir: "{app}"; Flags: ignoreversion signonce; Check: not Is64BitInstallMode
Source: "..\publish\ConformU86\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes:"*.exe,*.dll"; Check: not Is64BitInstallMode
Source: "..\publish\ConformU86\wwwroot\*"; DestDir: "{app}\wwwroot"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: not Is64BitInstallMode

[Icons]
Name: "{autoprograms}\ASCOM Conform Universal"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\ASCOM.ico"
Name: "{autodesktop}\Conform Universal"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\ASCOM.ico"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent unchecked

[UninstallDelete]
Name: "{app}\32bit"; Type: dirifempty
Name: "{app}"; Type: dirifempty

[Code]
procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpSelectTasks then
  begin
    WizardSelectTasks('windotnet');
  end;
end;

// Code to enable the installer to uninstall previous versions of itself when a new version is installed
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  UninstallExe: String;
  UninstallRegistry: String;
begin
  if (CurStep = ssInstall) then
	begin
      UninstallRegistry := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}' + '_is1');
      if RegQueryStringValue(HKLM, UninstallRegistry, 'UninstallString', UninstallExe) then
        begin
          Exec(RemoveQuotes(UninstallExe), ' /SILENT', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
          sleep(1000);    //Give enough time for the install screen to be repainted before continuing
        end
  end;
end;