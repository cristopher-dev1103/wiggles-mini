# =============================================================================
#  WIGGLES MINI - ARRANQUE REMOTO (estilo WinUtil)
#
#  Uso:
#    irm "https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1" | iex
#
#  Descarga la herramienta a %TEMP%\WigglesMini, pide permisos de
#  administrador y la abre. No instala nada de forma permanente.
#
#  Repositorio: https://github.com/cristopher-dev1103/wiggles-mini
# =============================================================================

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
#  CONFIGURACION DEL REPOSITORIO
# ---------------------------------------------------------------------------
$BaseUrl = 'https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main'

# Permite probar sin publicar:  $env:WIGGLES_BASEURL = 'http://127.0.0.1:8000'
if ($env:WIGGLES_BASEURL) { $BaseUrl = $env:WIGGLES_BASEURL.TrimEnd('/') }

$Destino = Join-Path $env:TEMP 'WigglesMini'
$Archivos = @(
    'WigglesMini.ps1'
    'WigglesCore.ps1'
    'config.json'
    'README.md'
    'Sheets\Code.gs'
    'External\Office\Configuration.xml'
    'External\Office\Setup.exe'
    'External\Tools\LEEME.txt'
)

# ---------------------------------------------------------------------------
#  DESCARGAR
# ---------------------------------------------------------------------------
function Escribir($Texto, $Color = 'Gray') {
    Write-Host $Texto -ForegroundColor $Color
}

function Descargar($Relativa) {
    $Url = $BaseUrl + '/' + ($Relativa -replace '\\', '/')
    $Ruta = Join-Path $Destino $Relativa

    $Carpeta = Split-Path $Ruta -Parent
    if (-not (Test-Path $Carpeta)) { New-Item -Path $Carpeta -ItemType Directory -Force | Out-Null }

    Invoke-WebRequest -Uri $Url -OutFile $Ruta -UseBasicParsing -TimeoutSec 90
    if (-not (Test-Path $Ruta)) { throw "No se descargo $Relativa" }
    if ((Get-Item $Ruta).Length -eq 0) { throw "Archivo vacio: $Relativa" }
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}

Escribir ''
Escribir '  WIGGLES MINI' 'Green'
Escribir '  Inventario para visitas a clientes' 'DarkGray'
Escribir ''
Escribir "  Descargando desde $BaseUrl" 'DarkGray'

if ($BaseUrl -match 'TU-USUARIO') {
    Escribir ''
    Escribir '  ERROR: falta configurar $BaseUrl con tu repositorio de GitHub.' 'Red'
    Escribir '  Edita irm.ps1 y cambia TU-USUARIO/wiggles-mini por el tuyo.' 'Yellow'
    return
}

$Inicio = Get-Date
try {
    foreach ($Archivo in $Archivos) {
        Escribir ("  - {0}" -f $Archivo) 'DarkGray'
        Descargar $Archivo
    }
} catch {
    Escribir ''
    Escribir "  ERROR al descargar: $($_.Exception.Message)" 'Red'
    Escribir '  Revisa que el repositorio sea publico y que la rama sea main.' 'Yellow'
    Escribir "  URL base: $BaseUrl" 'DarkGray'
    return
}

# ---------------------------------------------------------------------------
#  CONFIGURACION OPCIONAL POR VARIABLES DE ENTORNO
#  Util para probar sin publicar el token en el repositorio.
# ---------------------------------------------------------------------------
$ConfigPath = Join-Path $Destino 'config.json'
if ($env:WIGGLES_URL -or $env:WIGGLES_TOKEN) {
    try {
        $Cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($env:WIGGLES_URL) { $Cfg.WebhookUrl = $env:WIGGLES_URL }
        if ($env:WIGGLES_TOKEN) { $Cfg.Token = $env:WIGGLES_TOKEN }
        $Cfg | ConvertTo-Json -Depth 4 | Set-Content -Path $ConfigPath -Encoding UTF8
        Escribir ''
        Escribir '  Configuracion aplicada desde variables de entorno.' 'DarkGray'
    } catch {
        Escribir "  Aviso: no se pudo aplicar la configuracion: $($_.Exception.Message)" 'Yellow'
    }
}

try {
    $Cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace($Cfg.WebhookUrl)) {
        Escribir ''
        Escribir '  Aviso: config.json no tiene WebhookUrl.' 'Yellow'
        Escribir '  La herramienta funciona igual, pero los registros se quedan en el pendrive.' 'DarkGray'
    }
} catch {}

$Segundos = [math]::Round(((Get-Date) - $Inicio).TotalSeconds, 1)
Escribir ''
Escribir ("  Listo en {0} s." -f $Segundos) 'Green'

if ($env:WIGGLES_NO_LAUNCH -eq '1') {
    Escribir '  Modo prueba: no se abre la aplicacion.' 'Yellow'
    return
}

Escribir '  Abriendo como administrador...' 'Green'
Escribir ''

# ---------------------------------------------------------------------------
#  ABRIR ELEVADO
# ---------------------------------------------------------------------------
$Main = Join-Path $Destino 'WigglesMini.ps1'
$Argumentos = "-NoProfile -ExecutionPolicy Bypass -File `"$Main`""

$EsAdmin = $false
try {
    $Id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $EsAdmin = (New-Object Security.Principal.WindowsPrincipal($Id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {}

try {
    if ($EsAdmin) {
        Start-Process powershell.exe -ArgumentList $Argumentos -Wait
    } else {
        Start-Process powershell.exe -ArgumentList $Argumentos -Verb RunAs -Wait
    }
} catch {
    Escribir '  No se pudo abrir como administrador.' 'Red'
    Escribir '  Cancela el aviso de UAC y vuelve a intentarlo, o abre PowerShell como admin.' 'Yellow'
    Escribir "  Ruta: $Main" 'DarkGray'
}

Escribir ''
Escribir '  Sesion terminada. Los datos quedaron en:' 'DarkGray'
Escribir "  $Destino\Data" 'DarkGray'
Escribir ''
