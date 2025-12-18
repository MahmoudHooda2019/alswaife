; Inno Setup script for Al Sawife Factory Management
; 1) First run build_exe.bat to generate dist\AlSawifeFactory
; 2) Then open this .iss in Inno Setup and click Build

[Setup]
AppName=Al Sawife Factory Management
AppVersion=1.0.0
DefaultDirName={pf}\AlSawifeFactory
DefaultGroupName=Al Sawife Factory
OutputBaseFilename=AlSawifeFactorySetup
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=yes
SetupIconFile=res\icon.ico
UninstallDisplayIcon={app}\res\icon.ico

[Languages]
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Copy everything from PyInstaller output folder
Source: "dist\AlSawifeFactory\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Copy the icon file
Source: "res\icon.ico"; DestDir: "{app}\res"; Flags: ignoreversion

[Icons]
Name: "{group}\Al Sawife Factory"; Filename: "{app}\AlSawifeFactory.exe"; IconFilename: "{app}\res\icon.ico"
Name: "{commondesktop}\Al Sawife Factory"; Filename: "{app}\AlSawifeFactory.exe"; Tasks: desktopicon; IconFilename: "{app}\res\icon.ico"

[Run]
Filename: "{app}\AlSawifeFactory.exe"; Description: "تشغيل البرنامج الآن"; Flags: nowait postinstall skipifsilent