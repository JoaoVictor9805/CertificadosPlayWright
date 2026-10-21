; =====================================================================
; Script Inno Setup para o Robo DET (Certificados e Notificações)
; =====================================================================
#define MyAppName "Robo DET"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Societario"
#define MyAppExeName "RoboDET.exe"

[Setup]
AppId={{B426B97A-83F1-4F90-8AE2-68045F313171}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=Instalador_RoboDET
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Não exige privilégio de administrador (instala direto na pasta do usuário)
PrivilegesRequired=lowest
CloseApplications=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Copia todos os arquivos gerados pelo PyInstaller na pasta dist\RoboDET
Source: "dist\RoboDET\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
