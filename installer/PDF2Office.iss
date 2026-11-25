; ---------------------------------------------------------------------------
;  PDF2Office 安装包脚本（Inno Setup 6.5+ / 7.x）
;
;  编译：  powershell -ExecutionPolicy Bypass -File installer\build_installer.ps1
;  手动：  ISCC.exe /DAppVersion=2.0.1 installer\PDF2Office.iss
;
;  前置条件：先跑 build.ps1 生成本脚本要打包的 dist\PDF2Office
; ---------------------------------------------------------------------------

#define AppName        "PDF2Office"
#ifndef AppVersion
  #define AppVersion   "2.0.1"
#endif
#define AppVerName     AppName + " " + AppVersion
#define AppPublisher   "PDF2Office"
#define AppExeName     "PDF2Office.exe"
; 升级用的固定标识，发布后不要再改
#define AppId          "{{9A03D750-05C6-487B-8DCD-41727E0FE1B3}"
#define PayloadDir     "..\dist\PDF2Office"
#define IconFile       "..\assets\icon.ico"
#define ReadmeFile     "..\README.md"
#define DisclaimerFile "..\DISCLAIMER.md"
#define NoticeFile     "DISCLAIMER.txt"
#define InstallerName  AppName + "-" + AppVersion + "-Setup"
#define ReadmeName     "使用说明.md"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppVerName}
AppPublisher={#AppPublisher}
AppComments=PDF / Word / PPT 本地互相转换，离线运行
AppCopyright=© 2026 unclexiaoqi
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
DisableProgramGroupPage=yes
; 默认装到当前用户目录（不需要管理员、不弹 UAC）；安装向导里可切换到「所有用户」
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=output
OutputBaseFilename={#InstallerName}
SetupIconFile={#IconFile}
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppVerName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; 安装前先展示免责声明摘要（对外文档必须带的三处声明之一）
InfoBeforeFile={#NoticeFile}
; 自动按系统语言选择界面语言，不弹语言选择框（中文系统→简体中文，其它→英文）
ShowLanguageDialog=no
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
CloseApplications=yes
RestartApplications=no
VersionInfoVersion={#AppVersion}.0
VersionInfoCompany={#AppPublisher}
VersionInfoCopyright=© 2026 unclexiaoqi
VersionInfoDescription={#AppVerName} 安装程序
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}

[Languages]
; 语言文件随仓库自带：Inno Setup 6.7+ 的安装包已不再附带非官方翻译（简体中文就在其中）
Name: "chinesesimplified"; MessagesFile: "languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; 整个 PyInstaller 产物（PDF2Office.exe + _internal）原样安装
Source: "{#PayloadDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#ReadmeFile}"; DestDir: "{app}"; DestName: "{#ReadmeName}"; Flags: ignoreversion
Source: "{#DisclaimerFile}"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; 让「Win+R → PDF2Office」和命令行直接可用（管理员装写 HKLM，普通装写 HKCU）
Root: HKA; Subkey: "Software\Microsoft\Windows\CurrentVersion\App Paths\{#AppExeName}"; ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\使用说明"; Filename: "{app}\{#ReadmeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[CustomMessages]
chinesesimplified.RemoveSettingsPrompt=是否同时删除 PDF2Office 的设置和错误日志？%n%n%1
english.RemoveSettingsPrompt=Also delete the PDF2Office settings and error log?%n%n%1

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  SettingsDir: String;
begin
  // 提示是否一并删除 %APPDATA%\PDF2Office（静默卸载时保留，避免脚本误删用户数据）
  if CurUninstallStep = usPostUninstall then
  begin
    SettingsDir := ExpandConstant('{userappdata}\PDF2Office');
    if DirExists(SettingsDir) and not UninstallSilent then
    begin
      // 默认按钮是「否」，回车不会误删
      if MsgBox(FmtMessage(CustomMessage('RemoveSettingsPrompt'), [SettingsDir]),
                mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
        DelTree(SettingsDir, True, True, True);
    end;
  end;
end;
