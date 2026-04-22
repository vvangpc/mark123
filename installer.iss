; installer.iss — Inno Setup 安装脚本
; 用法：
;   1. 先跑 .venv\Scripts\python.exe build.py onedir   生成 dist\专利标记助手\
;   2. 安装 Inno Setup 6 (https://jrsoftware.org/isdl.php)
;   3. 在 Inno Setup Compiler 里打开本文件 → Build → Compile
;      或命令行：iscc installer.iss
;   4. 产物：dist\专利标记助手V3.5-安装版.exe

#define MyAppName "专利标记助手"
#define MyAppVersion "3.5"
#define MyAppPublisher "vvangpc"
#define MyAppExeName "专利标记助手.exe"
#define MyAppURL "https://github.com/vvangpc/mark123"

[Setup]
AppId={{8F3D5E2A-4B7C-4D6A-9E3B-7A2F6C1D8E5B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=dist
OutputBaseFilename=专利标记助手V{#MyAppVersion}-安装版
Compression=lzma2/ultra
SolidCompression=yes
WizardStyle=modern
; 默认按管理员权限安装到 Program Files。也可改为 lowest 装到用户目录。
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; 简体中文界面
LanguageDetectionMethod=none
ShowLanguageDialog=no

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加快捷方式:"
Name: "docxcontextmenu"; Description: "在 .docx 文件的右键菜单中添加「用专利标记助手打开」"; GroupDescription: "文件关联:"; Flags: checkedonce

[Files]
; 把 onedir 产出的整个文件夹递归塞进 {app}
Source: "dist\专利标记助手\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; 在所有 .docx 文件的右键菜单里加一项「用专利标记助手打开」
; 走 SystemFileAssociations，不影响 Word 的默认打开行为，纯粹新增一个菜单项
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\{#MyAppName}"; \
    ValueType: string; ValueName: ""; ValueData: "用{#MyAppName}打开"; \
    Flags: uninsdeletekey; Tasks: docxcontextmenu
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\{#MyAppName}"; \
    ValueType: string; ValueName: "Icon"; ValueData: """{app}\{#MyAppExeName}"""; \
    Tasks: docxcontextmenu
Root: HKA; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\{#MyAppName}\command"; \
    ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; \
    Flags: uninsdeletekey; Tasks: docxcontextmenu

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName}"; \
    Flags: nowait postinstall skipifsilent
