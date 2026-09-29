# Instala o Redator para o usuário atual, sem precisar de administrador.
# Copia o programa, cria atalhos e registra os tipos de arquivo no "Abrir com".
$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot '..\out\Redator-win32-x64'
$dest = Join-Path $env:LOCALAPPDATA 'Programs\Redator'
$exe = Join-Path $dest 'Redator.exe'

# fecha instâncias abertas para poder substituir os arquivos
Get-Process Redator -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
Start-Sleep -Milliseconds 800
Get-Process Redator -ErrorAction SilentlyContinue | Stop-Process -Force

New-Item -ItemType Directory -Force -Path $dest | Out-Null
robocopy $src $dest /MIR /NFL /NDL /NJH /NJS /NP | Out-Null

# atalhos
$shell = New-Object -ComObject WScript.Shell
$startMenu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Redator.lnk'
$desktop = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Redator.lnk'
foreach ($lnkPath in @($startMenu, $desktop)) {
  $lnk = $shell.CreateShortcut($lnkPath)
  $lnk.TargetPath = $exe
  $lnk.WorkingDirectory = $dest
  $lnk.IconLocation = "$exe,0"
  $lnk.Description = 'Redator - editor de documentos'
  $lnk.Save()
}

# tipos de arquivo
$classes = 'HKCU:\Software\Classes'
$types = [ordered]@{
  '.docx' = 'Documento do Word'
  '.docm' = 'Documento do Word Habilitado para Macro'
  '.dotx' = 'Modelo do Word'
  '.odt' = 'Texto OpenDocument'
  '.md' = 'Documento Markdown'
  '.markdown' = 'Documento Markdown'
  '.txt' = 'Documento de Texto'
  '.rtf' = 'Documento Rich Text'
  '.doc' = 'Documento do Word 97-2003'
  '.pdf' = 'Documento PDF'
  '.dotm' = 'Modelo do Word Habilitado para Macro'
  '.ott' = 'Modelo OpenDocument'
  '.fodt' = 'Texto OpenDocument plano'
  '.epub' = 'Livro digital EPUB'
  '.mht' = 'Página da Web em arquivo único'
  '.mhtml' = 'Página da Web em arquivo único'
  '.eml' = 'Mensagem de e-mail'
  '.csv' = 'Arquivo CSV'
  '.tsv' = 'Valores separados por tabulação'
  '.log' = 'Arquivo de log'
  '.json' = 'Arquivo JSON'
  '.xml' = 'Documento XML'
  '.yaml' = 'Arquivo YAML'
  '.yml' = 'Arquivo YAML'
  '.ini' = 'Arquivo de configuração'
}
function Set-Default($path, $value) {
  if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
  Set-Item -Path $path -Value $value
}
foreach ($ext in $types.Keys) {
  $progId = 'Redator' + $ext
  Set-Default "$classes\$progId" $types[$ext]
  Set-Default "$classes\$progId\DefaultIcon" "`"$exe`",0"
  Set-Default "$classes\$progId\shell\open\command" "`"$exe`" `"%1`""
  if (-not (Test-Path "$classes\$ext\OpenWithProgids")) { New-Item -Path "$classes\$ext\OpenWithProgids" -Force | Out-Null }
  New-ItemProperty -Path "$classes\$ext\OpenWithProgids" -Name $progId -Value '' -PropertyType String -Force | Out-Null
}
Set-Default "$classes\Applications\Redator.exe\shell\open\command" "`"$exe`" `"%1`""
New-ItemProperty -Path "$classes\Applications\Redator.exe" -Name 'FriendlyAppName' -Value 'Redator' -PropertyType String -Force | Out-Null
if (-not (Test-Path "$classes\Applications\Redator.exe\SupportedTypes")) { New-Item -Path "$classes\Applications\Redator.exe\SupportedTypes" -Force | Out-Null }
foreach ($ext in $types.Keys) { New-ItemProperty -Path "$classes\Applications\Redator.exe\SupportedTypes" -Name $ext -Value '' -PropertyType String -Force | Out-Null }

# capacidades, para aparecer em Configurações > Aplicativos padrão
$cap = 'HKCU:\Software\Redator\Capabilities'
if (-not (Test-Path "$cap\FileAssociations")) { New-Item -Path "$cap\FileAssociations" -Force | Out-Null }
New-ItemProperty -Path $cap -Name 'ApplicationName' -Value 'Redator' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $cap -Name 'ApplicationDescription' -Value 'Editor de documentos com a cara do Word: docx, doc, pdf, odt, rtf, md e txt' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $cap -Name 'ApplicationIcon' -Value "`"$exe`",0" -PropertyType String -Force | Out-Null
foreach ($ext in $types.Keys) { New-ItemProperty -Path "$cap\FileAssociations" -Name $ext -Value ('Redator' + $ext) -PropertyType String -Force | Out-Null }
if (-not (Test-Path 'HKCU:\Software\RegisteredApplications')) { New-Item -Path 'HKCU:\Software\RegisteredApplications' -Force | Out-Null }
New-ItemProperty -Path 'HKCU:\Software\RegisteredApplications' -Name 'Redator' -Value 'Software\Redator\Capabilities' -PropertyType String -Force | Out-Null

# avisa o Explorer que as associações mudaram
Add-Type -Namespace Win32 -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);'
[Win32.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)

"Redator instalado em $dest"
