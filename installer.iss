[Setup]
AppName=Modern Drive Search
AppVersion=1.0
DefaultDirName={pf}\ModernDriveSearch
DefaultGroupName=ModernDriveSearch
OutputDir=Output
OutputBaseFilename=ModernDriveSearchSetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\ModernDriveSearch.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Modern Drive Search"; Filename: "{app}\ModernDriveSearch.exe"
Name: "{commondesktop}\Modern Drive Search"; Filename: "{app}\ModernDriveSearch.exe"