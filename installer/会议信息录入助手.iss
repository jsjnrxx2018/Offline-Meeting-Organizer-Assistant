[Setup]
AppId={{B8B6DB6E-7C75-4B2D-9F6C-7C726AF65E3A}
AppName=会议信息录入助手
AppVersion=1.0.0
AppPublisher=上海第二工业大学
DefaultDirName={autopf}\会议信息录入助手
DefaultGroupName=会议信息录入助手
OutputDir=.
OutputBaseFilename=会议信息录入助手-安装包
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加图标："; Flags: unchecked

[Files]
Source: "..\dist\会议信息录入助手\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\会议信息录入助手"; Filename: "{app}\会议信息录入助手.exe"
Name: "{group}\卸载会议信息录入助手"; Filename: "{uninstallexe}"
Name: "{autodesktop}\会议信息录入助手"; Filename: "{app}\会议信息录入助手.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\会议信息录入助手.exe"; Description: "启动会议信息录入助手"; Flags: nowait postinstall skipifsilent
