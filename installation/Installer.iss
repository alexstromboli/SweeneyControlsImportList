[Setup]
AppId={{3A24D3BC-603E-4ECA-92F0-AC8EF3C86AA1}
AppName=SweeneyControls ImportList
AppVersion=2.0
UninstallDisplayName=SweeneyControls ImportList
DefaultDirName={commonappdata}\Autodesk\ApplicationPlugins\SweeneyControls.bundle
DisableDirPage=yes
DefaultGroupName=SweeneyControls
DisableProgramGroupPage=yes
OutputBaseFilename=install
OutputDir=output
Compression=lzma
SolidCompression=yes

[Files]
Source: "..\src\ImportList\bin\Debug\ImportList.dll"; DestDir: "{app}\Contents"; Flags: ignoreversion
Source: "..\src\ImportList\bin\Debug\SpreadsheetGear.dll"; DestDir: "{app}\Contents"; Flags: ignoreversion
Source: "PackageContents.xml"; DestDir: "{app}"; Flags: ignoreversion
