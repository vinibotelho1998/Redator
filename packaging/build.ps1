# Gera o instalador do Redator a partir de um clone limpo.
#
# Usa o Node que encontrar, instala as dependencias se faltarem, empacota o
# Electron e entrega um unico arquivo em out\Redator-Setup.exe, que e' o que se
# manda para outra pessoa.
#
#   powershell -ExecutionPolicy Bypass -File packaging\build.ps1
#
# -SemInstalador para so' gerar a pasta out\Redator-win32-x64 (util para testar rapido).
#
# Sem acentos de proposito: o PowerShell 5.1 le .ps1 como ANSI e um caractere
# fora do ASCII quebra o parser inteiro na maquina de quem for compilar.
param([switch]$SemInstalador)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Passo($txt) {
    Write-Host ""
    Write-Host "== $txt" -ForegroundColor Cyan
}
function Ok($txt)   { Write-Host "   ok    $txt" -ForegroundColor Green }
function Info($txt) { Write-Host "   ..    $txt" -ForegroundColor DarkGray }

# -------------------------------------------------------------------- Node
# Algumas dependencias do empacotamento pedem Node 22.12 ou mais novo.
Passo "Procurando Node 22 ou mais novo"
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) { throw "Node nao encontrado. Instale o Node 22 LTS ou mais novo de nodejs.org e rode de novo." }
$versaoNode = (& node --version).Trim()
if ($versaoNode -notmatch '^v(\d+)\.' -or [int]$Matches[1] -lt 22) { throw "Node $versaoNode e' antigo demais. Instale o 22 LTS ou mais novo." }
Ok "Node $versaoNode"

# ------------------------------------------------------------ Dependencias
Push-Location $root
try {
    Passo "Preparando dependencias"
    if (-not (Test-Path (Join-Path $root "node_modules\electron\package.json"))) {
        Info "instalando pacotes do npm"
        & npm ci
        if ($LASTEXITCODE -ne 0) { throw "npm ci falhou (codigo $LASTEXITCODE)" }
    }
    Ok "dependencias prontas"

    $pkg = Get-Content (Join-Path $root "package.json") -Raw | ConvertFrom-Json
    $versao = $pkg.version
    Ok "versao do app: $versao"

    # --------------------------------------------------------------- Empacotar
    Passo "Empacotando com o Electron"
    & npm run package
    if ($LASTEXITCODE -ne 0) { throw "Empacotamento falhou (codigo $LASTEXITCODE)" }
} finally {
    Pop-Location
}
$appDir = Join-Path $root "out\Redator-win32-x64"
if (-not (Test-Path (Join-Path $appDir "Redator.exe"))) { throw "Redator.exe nao foi gerado" }
Ok "app em out\Redator-win32-x64"

if ($SemInstalador) {
    Write-Host ""
    Write-Host "Pronto (sem instalador): $appDir" -ForegroundColor Green
    return
}

# -------------------------------------------------------------- Inno Setup
# O link direto do site do Inno Setup passa por uma pagina intermediaria e
# devolve HTML no lugar do instalador. O winget resolve isso e ja' vem no
# Windows 10/11, entao e' por ele que a instalacao acontece.
function Find-Iscc {
    @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}

Passo "Procurando o Inno Setup"
$iscc = Find-Iscc
if (-not $iscc) {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Info "nao encontrado; instalando via winget"
        & winget install --id JRSoftware.InnoSetup -e --silent `
            --accept-source-agreements --accept-package-agreements | Out-Null
        $iscc = Find-Iscc
    }
}
if (-not $iscc) {
    throw ("Inno Setup 6 nao encontrado. Instale de https://jrsoftware.org/isdl.php " +
           "e rode de novo, ou use 'packaging\build.ps1 -SemInstalador' para gerar " +
           "so' a pasta out\Redator-win32-x64.")
}
Ok (Split-Path $iscc -Parent)

# -------------------------------------------------------------- Instalador
Passo "Compilando o instalador"
& $iscc "/DMyAppVersion=$versao" (Join-Path $PSScriptRoot "Redator_Setup.iss")
if ($LASTEXITCODE -ne 0) { throw "Inno Setup falhou (codigo $LASTEXITCODE)" }

$saida = Join-Path $root "out\Redator-Setup.exe"
$tamanho = [math]::Round((Get-Item $saida).Length / 1MB, 1)
Write-Host ""
Write-Host "=====================================================" -ForegroundColor DarkCyan
Write-Host " INSTALADOR PRONTO - versao $versao ($tamanho MB)" -ForegroundColor Green
Write-Host " $saida" -ForegroundColor White
Write-Host "=====================================================" -ForegroundColor DarkCyan
Write-Host " Esse arquivo e' tudo que a outra pessoa precisa." -ForegroundColor Gray
Write-Host " Nao exige Node, nem o repositorio, nem configurar nada." -ForegroundColor Gray