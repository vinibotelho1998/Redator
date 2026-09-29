; Instalador do Redator.
;
; Compile por packaging/build.ps1 (ou BUILD.bat): e' de la' que vem a versao,
; lida do package.json, para nao existir numero repetido em dois lugares.
; Compilar este arquivo na mao tambem funciona, usando a reserva.

#define MyAppName "Redator"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "Vinicius Botelho"
#define MyAppExeName "Redator.exe"
#define MyAppUrl "https://github.com/vinibotelho1998/Redator"

[Setup]
; AppId fixo e unico: e' por ele que o Windows reconhece uma atualizacao no
; lugar de uma segunda instalacao. Nunca mudar entre versoes.
AppId={{1501E501-4F6B-47E3-B2D0-BDD19FE5BAC6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppUrl}
AppSupportURL={#MyAppUrl}/issues
AppUpdatesURL={#MyAppUrl}/releases/latest
AppCopyright=Copyright (C) {#MyAppPublisher}
VersionInfoVersion={#MyAppVersion}
VersionInfoProductName={#MyAppName}
VersionInfoDescription=Instalador do {#MyAppName}
AppComments=Editor de documentos com a cara do Word: docx, odt, md, txt e PDF.
VersionInfoCompany={#MyAppPublisher}

DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableDirPage=auto

OutputDir=..\out
OutputBaseFilename=Redator-Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\app\icons\redator.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName} {#MyAppVersion}

; O editor nao mexe em nada do sistema: instala para o usuario atual e nao
; pede UAC. Quem rodar o setup como admin instala para a maquina inteira.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
; Fecha o Redator antes de sobrescrever os arquivos, em vez de falhar no meio.
CloseApplications=yes
RestartApplications=no
SetupMutex=RedatorSetupMutex
; Avisa o Explorer que os tipos de arquivo mudaram.
ChangesAssociations=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"
Name: "associar"; Description: "Oferecer o Redator no ""Abrir com"" de .docx, .doc, .odt, .md, .rtf e .txt"; GroupDescription: "Tipos de arquivo:"

[Files]
Source: "..\out\Redator-win32-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; O Windows nao deixa um programa se declarar padrao sozinho: o instalador
; registra os tipos no "Abrir com" e em Aplicativos padrao, e a escolha final
; fica com quem usa, uma vez por extensao.
Root: HKA; Subkey: "Software\Classes\Redator.docx"; ValueType: string; ValueName: ""; ValueData: "Documento do Word"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.docx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.docx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.docx\OpenWithProgids"; ValueType: string; ValueName: "Redator.docx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.docm"; ValueType: string; ValueName: ""; ValueData: "Documento do Word Habilitado para Macro"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.docm\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.docm\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.docm\OpenWithProgids"; ValueType: string; ValueName: "Redator.docm"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.dotx"; ValueType: string; ValueName: ""; ValueData: "Modelo do Word"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.dotx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.dotx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.dotx\OpenWithProgids"; ValueType: string; ValueName: "Redator.dotx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.odt"; ValueType: string; ValueName: ""; ValueData: "Texto OpenDocument"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.odt\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.odt\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.odt\OpenWithProgids"; ValueType: string; ValueName: "Redator.odt"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.md"; ValueType: string; ValueName: ""; ValueData: "Documento Markdown"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.md\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.md\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.md\OpenWithProgids"; ValueType: string; ValueName: "Redator.md"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.markdown"; ValueType: string; ValueName: ""; ValueData: "Documento Markdown"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.markdown\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.markdown\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.markdown\OpenWithProgids"; ValueType: string; ValueName: "Redator.markdown"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.txt"; ValueType: string; ValueName: ""; ValueData: "Documento de Texto"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.txt\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.txt\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.txt\OpenWithProgids"; ValueType: string; ValueName: "Redator.txt"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.rtf"; ValueType: string; ValueName: ""; ValueData: "Documento Rich Text"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.rtf\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.rtf\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.rtf\OpenWithProgids"; ValueType: string; ValueName: "Redator.rtf"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.doc"; ValueType: string; ValueName: ""; ValueData: "Documento do Word 97-2003"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.doc\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Redator.doc\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associar
Root: HKA; Subkey: "Software\Classes\.doc\OpenWithProgids"; ValueType: string; ValueName: "Redator.doc"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associar
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".docx"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".docm"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".dotx"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".odt"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".md"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".markdown"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".txt"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".rtf"; ValueData: ""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".doc"; ValueData: ""
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "{#MyAppName}"; Flags: uninsdeletekey; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "Editor de documentos com a cara do Word: docx, odt, md, txt e PDF"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities"; ValueType: string; ValueName: "ApplicationIcon"; ValueData: """{app}\{#MyAppExeName}"",0"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".docx"; ValueData: "Redator.docx"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".docm"; ValueData: "Redator.docm"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".dotx"; ValueData: "Redator.dotx"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".odt"; ValueData: "Redator.odt"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".md"; ValueData: "Redator.md"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".markdown"; ValueData: "Redator.markdown"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".txt"; ValueData: "Redator.txt"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".rtf"; ValueData: "Redator.rtf"; Tasks: associar
Root: HKA; Subkey: "Software\{#MyAppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".doc"; ValueData: "Redator.doc"; Tasks: associar
Root: HKA; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: "Software\{#MyAppName}\Capabilities"; Flags: uninsdeletevalue; Tasks: associar

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir o {#MyAppName} agora"; Flags: nowait postinstall skipifsilent

[Messages]
FinishedLabel=A instalação do [name] terminou.%n%nPara abrir documentos com ele por padrão, clique com o botão direito em um arquivo .docx ou .md, escolha Abrir com, Escolher outro aplicativo, marque o Redator e Sempre.

[Code]
// Preferencias, arquivos recentes e o dicionario pessoal moram em
// %APPDATA%\Redator. Apagar isso sem perguntar seria perder as palavras
// que a pessoa adicionou ao corretor.
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Dados: string;
begin
  if CurUninstallStep <> usPostUninstall then
    Exit;
  Dados := ExpandConstant('{userappdata}\Redator');
  if not DirExists(Dados) then
    Exit;
  if MsgBox('Apagar também as preferências, a lista de arquivos recentes e as palavras adicionadas ao dicionário?'
            + #13#10 + #13#10 + Dados,
            mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
    DelTree(Dados, True, True, True);
end;