# =============================================================================
#  WIGGLES MINI - CORE
#  Escaneo de hardware, exportacion y sincronizacion con Google Sheets.
#  Sin dependencias externas. PowerShell 5.1+ / Windows 10-11.
# =============================================================================

$script:Culture = [System.Globalization.CultureInfo]::InvariantCulture

# Determina si el proceso esta elevado (TPM y SecureBoot requieren admin)
$script:EsAdmin = $false
try {
    $Identidad = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($Identidad)
    $script:EsAdmin = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {}

# -----------------------------------------------------------------------------
#  ORDEN DE COLUMNAS (una sola fuente de verdad para CSV, Excel y Sheets)
# -----------------------------------------------------------------------------
$script:Columnas = @(
    'Cliente'
    'Sede_Tecnico'
    'Orden'
    'Fecha'
    'Hora'
    'Hostname'
    'Usuario'
    'Tipo_Equipo'
    'Marca'
    'Modelo'
    'Serial'
    'Serial_Chasis'
    'Asset_Tag'
    'CPU_Modelo'
    'CPU_Nucleos'
    'CPU_Hilos'
    'CPU_GHz'
    'RAM_Total_GB'
    'RAM_Tipo'
    'RAM_Velocidad_MHz'
    'RAM_Slots'
    'RAM_Slots_Libres'
    'RAM_Max_GB'
    'RAM_Modulos'
    'GPU'
    'GPU_VRAM'
    'Placa_Madre'
    'BIOS_Version'
    'BIOS_Fecha'
    'Monitores'
    'Resolucion'
    'Disco_Modelo'
    'Disco_Tipo'
    'Disco_Capacidad_GB'
    'Disco_Libre_GB'
    'Disco_Uso_Pct'
    'Disco_Salud'
    'Disco_Temperatura_C'
    'Disco_Horas_Encendido'
    'Discos_Extra'
    'Bateria_Presente'
    'Bateria_Diseno_mWh'
    'Bateria_Actual_mWh'
    'Bateria_Desgaste_Pct'
    'Bateria_Ciclos'
    'Bateria_Estado'
    'Bateria_Voltaje_V'
    'Bateria_Corriente_A'
    'Bateria_Potencia_W'
    'SO_Nombre'
    'SO_Version'
    'SO_Build'
    'SO_Arquitectura'
    'SO_Instalado'
    'Ultimo_Arranque'
    'Uptime_Horas'
    'Windows_Activado'
    'Office_Version'
    'Office_Activado'
    'Dominio_Grupo'
    'IP'
    'MAC'
    'Gateway'
    'Adaptadores'
    'TPM'
    'SecureBoot'
    'Antivirus'
    'Software_Total'
    'Software_Lista'
    'Comentarios'
)

# -----------------------------------------------------------------------------
#  UTILIDADES
# -----------------------------------------------------------------------------
function ConvertTo-Plano {
    param($Valor, [int]$Max = 30000)
    if ($null -eq $Valor) { return '' }
    $t = [string]$Valor
    $t = $t -replace '\s+', ' '
    $t = $t.Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max) + '...' }
    return $t
}

function Get-CimSeguro {
    param(
        [string]$Clase,
        [string]$Namespace = 'root/cimv2',
        [string]$Filtro
    )
    try {
        if ($Filtro) {
            return Get-CimInstance -ClassName $Clase -Namespace $Namespace -Filter $Filtro -ErrorAction Stop
        }
        return Get-CimInstance -ClassName $Clase -Namespace $Namespace -ErrorAction Stop
    } catch {
        return $null
    }
}

function ConvertTo-GB {
    param($Bytes, [int]$Decimales = 0)
    if ($null -eq $Bytes -or $Bytes -eq 0) { return 0 }
    return [math]::Round(([double]$Bytes / 1GB), $Decimales)
}

# -----------------------------------------------------------------------------
#  RAM
# -----------------------------------------------------------------------------
function Get-RAMDetalle {
    $Resultado = [ordered]@{
        TotalGB      = 0
        Tipo         = 'N/A'
        Velocidad    = 'N/A'
        SlotsUsados  = 0
        SlotsTotales = 0
        SlotsLibres  = 0
        MaxGB        = 0
        Modulos      = 'N/A'
    }

    $Modulos = Get-CimSeguro -Clase 'Win32_PhysicalMemory'
    if ($Modulos) {
        $TotalBytes = 0
        $Partes = @()
        $Velocidades = @()
        $TiposCodigo = @()

        foreach ($m in $Modulos) {
            $Cap = 0
            try { $Cap = [double]$m.Capacity } catch {}
            $TotalBytes += $Cap
            $Partes += ('{0} {1}GB {2}' -f $m.DeviceLocator, (ConvertTo-GB $Cap), $m.Manufacturer).Trim()
            if ($m.ConfiguredClockSpeed -gt 0) { $Velocidades += [int]$m.ConfiguredClockSpeed }
            elseif ($m.Speed -gt 0) { $Velocidades += [int]$m.Speed }
            if ($m.SMBIOSMemoryType) { $TiposCodigo += [int]$m.SMBIOSMemoryType }
        }

        $Resultado.TotalGB = ConvertTo-GB $TotalBytes
        $Resultado.SlotsUsados = @($Modulos).Count

        if ($Velocidades.Count -gt 0) {
            $Resultado.Velocidad = ('{0} MHz' -f (($Velocidades | Measure-Object -Maximum).Maximum))
        }

        if ($TiposCodigo.Count -gt 0) {
            $MapaTipo = @{
                20 = 'DDR'; 21 = 'DDR2'; 22 = 'DDR2 FB-DIMM'; 24 = 'DDR3';
                26 = 'DDR4'; 27 = 'LPDDR'; 28 = 'LPDDR2'; 29 = 'LPDDR3';
                30 = 'LPDDR4'; 34 = 'DDR5'; 35 = 'LPDDR5'
            }
            $Cod = ($TiposCodigo | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
            $Resultado.Tipo = if ($MapaTipo.ContainsKey([int]$Cod)) { $MapaTipo[[int]$Cod] } else { "Codigo $Cod" }
        }

        $Resultado.Modulos = ConvertTo-Plano ($Partes -join ' | ') 2000
    }

    $Array = Get-CimSeguro -Clase 'Win32_PhysicalMemoryArray'
    if ($Array) {
        $Arr = @($Array)[0]
        try { $Resultado.SlotsTotales = [int]$Arr.MemoryDevices } catch {}
        try {
            # MaxCapacity y MaxCapacityEx vienen en KB segun SMBIOS
            if ($Arr.MaxCapacityEx -gt 0) { $Resultado.MaxGB = ConvertTo-GB ([double]$Arr.MaxCapacityEx * 1KB) }
            elseif ($Arr.MaxCapacity -gt 0) { $Resultado.MaxGB = ConvertTo-GB ([double]$Arr.MaxCapacity * 1KB) }
        } catch {
            try { if ($Arr.MaxCapacity -gt 0) { $Resultado.MaxGB = ConvertTo-GB ([double]$Arr.MaxCapacity * 1KB) } } catch {}
        }
    }

    if ($Resultado.SlotsTotales -eq 0) { $Resultado.SlotsTotales = $Resultado.SlotsUsados }
    $Resultado.SlotsLibres = [math]::Max(0, $Resultado.SlotsTotales - $Resultado.SlotsUsados)

    return [PSCustomObject]$Resultado
}

# -----------------------------------------------------------------------------
#  DISCO / SMART
# -----------------------------------------------------------------------------
function Get-DiscoDetalle {
    $Resultado = [ordered]@{
        Modelo        = 'N/A'
        Tipo          = 'N/A'
        CapacidadGB   = 0
        LibreGB       = 0
        UsoPct        = ''
        Salud         = 'N/A'
        Temperatura   = ''
        Horas         = ''
        Extra         = ''
    }

    # Discos fisicos (Storage module: Win8+)
    $Fisicos = $null
    try { $Fisicos = Get-PhysicalDisk -ErrorAction Stop } catch {}

    if ($Fisicos) {
        # Excluir lectores de tarjetas / unidades USB (reportan el mismo tamano y ensucian el inventario)
        $Lista = @($Fisicos) | Where-Object { [string]$_.BusType -ne 'USB' -and $_.Size -gt 0 }
        if ($Lista.Count -eq 0) { $Lista = @($Fisicos) | Where-Object { $_.Size -gt 0 } }
        $Principal = $Lista | Sort-Object Size -Descending | Select-Object -First 1

        $Resultado.Modelo = ConvertTo-Plano $Principal.FriendlyName 120
        $Resultado.CapacidadGB = ConvertTo-GB $Principal.Size

        # MediaType / BusType llegan como enum en texto ("SSD", "NVMe") o numero
        $MediaTxt = [string]$Principal.MediaType
        $TipoMedia = switch -Regex ($MediaTxt) {
            '^3$|HDD'      { 'HDD' }
            '^4$|SSD'      { 'SSD' }
            '^5$|SCM'      { 'SCM' }
            default        { 'Desconocido' }
        }
        $BusTxt = [string]$Principal.BusType
        $Bus = switch -Regex ($BusTxt) {
            '^7$|USB'      { 'USB' }
            '^8$|RAID'     { 'RAID' }
            '^11$|SATA'    { 'SATA' }
            '^10$|SAS'     { 'SAS' }
            '^17$|NVMe'    { 'NVMe' }
            '^14$|Virtual' { 'Virtual' }
            default        { '' }
        }
        $Resultado.Tipo = (($TipoMedia, $Bus | Where-Object { $_ }) -join ' ')
        $Resultado.Salud = [string]$Principal.HealthStatus

        # SMART (requiere admin, no siempre disponible)
        try {
            $Smart = $Principal | Get-StorageReliabilityCounter -ErrorAction Stop
            if ($Smart) {
                if ($null -ne $Smart.Temperature -and $Smart.Temperature -gt 0) { $Resultado.Temperatura = [string]$Smart.Temperature }
                if ($null -ne $Smart.PowerOnHours -and $Smart.PowerOnHours -gt 0) { $Resultado.Horas = [string]$Smart.PowerOnHours }
                if ($null -ne $Smart.Wear -and $Smart.Wear -gt 0) {
                    $Resultado.Salud = ('{0} (desgaste {1}%)' -f $Resultado.Salud, $Smart.Wear)
                }
            }
        } catch {}

        if ($Lista.Count -gt 1) {
            $Extra = $Lista | Sort-Object Size -Descending | Select-Object -Skip 1 | ForEach-Object {
                '{0} ({1}GB)' -f $_.FriendlyName, (ConvertTo-GB $_.Size)
            }
            $Resultado.Extra = ConvertTo-Plano ($Extra -join ' | ') 500
        }
    } else {
        # Fallback WMI clasico
        $DD = Get-CimSeguro -Clase 'Win32_DiskDrive'
        if ($DD) {
            $Principal = @($DD) | Sort-Object Size -Descending | Select-Object -First 1
            $Resultado.Modelo = ConvertTo-Plano $Principal.Model 120
            $Resultado.CapacidadGB = ConvertTo-GB $Principal.Size
            $Resultado.Tipo = if ($Principal.MediaType -match 'SSD') { 'SSD' } elseif ($Principal.MediaType -match 'Fixed') { 'HDD' } else { [string]$Principal.MediaType }
            $Resultado.Salud = [string]$Principal.Status
        }
    }

    $Sistema = $env:SystemDrive
    if (-not $Sistema) { $Sistema = 'C:' }
    $Vol = Get-CimSeguro -Clase 'Win32_LogicalDisk' -Filtro ("DeviceID='{0}'" -f $Sistema)
    if ($Vol) {
        $Resultado.LibreGB = ConvertTo-GB $Vol.FreeSpace
        if ($Vol.Size -gt 0) {
            $Resultado.UsoPct = [string]([math]::Round((($Vol.Size - $Vol.FreeSpace) / $Vol.Size) * 100, 0))
        }
    }

    return [PSCustomObject]$Resultado
}

# -----------------------------------------------------------------------------
#  BATERIA
# -----------------------------------------------------------------------------
function Get-BateriaDetalle {
    $Resultado = [ordered]@{
        Presente   = 'No'
        Diseno     = ''
        Actual     = ''
        Desgaste   = ''
        Ciclos     = ''
        Estado     = ''
        Voltaje    = ''
        Amperaje   = ''
        Watts      = ''
    }

    $Bat = Get-CimSeguro -Clase 'Win32_Battery'
    if (-not $Bat) { return [PSCustomObject]$Resultado }
    $B = @($Bat)[0]
    $Resultado.Presente = 'Si'
    $Resultado.Estado = ConvertTo-Plano $B.Name 80

    # --- Capacidad de diseno y ciclos: powercfg es rapido (~0.1 s) y confiable ---
    $Tmp = Join-Path $env:TEMP ('wiggles_bat_' + [guid]::NewGuid().ToString('N') + '.xml')
    try {
        powercfg /batteryreport /XML /OUTPUT "$Tmp" | Out-Null
        if (Test-Path $Tmp) {
            [xml]$Xml = Get-Content -Path $Tmp -Raw -ErrorAction Stop
            $Info = $Xml.BatteryReport.Batteries.Battery
            if ($Info) {
                $Info = @($Info)[0]
                if ($Info.DesignCapacity -gt 0)     { $Resultado.Diseno = [string][int]$Info.DesignCapacity }
                if ($Info.FullChargeCapacity -gt 0) { $Resultado.Actual = [string][int]$Info.FullChargeCapacity }
                if ($null -ne $Info.CycleCount)     { $Resultado.Ciclos = [string][int]$Info.CycleCount }
            }
        }
    } catch {
    } finally {
        # [IO.File]::Delete evita el bug de Remove-Item con rutas cortas 8.3 (C:\Users\XXXX~1)
        try { if ([System.IO.File]::Exists($Tmp)) { [System.IO.File]::Delete($Tmp) } } catch {}
    }

    # Fallback: clases WMI de bateria
    if (-not $Resultado.Actual) {
        $Cargada = Get-CimSeguro -Clase 'BatteryFullChargedCapacity' -Namespace 'root/wmi'
        if ($Cargada) {
            try { if (@($Cargada)[0].FullChargedCapacity -gt 0) { $Resultado.Actual = [string][int]@($Cargada)[0].FullChargedCapacity } } catch {}
        }
    }

    if (-not $Resultado.Diseno) {
        $Estatica = Get-CimSeguro -Clase 'BatteryStaticData' -Namespace 'root/wmi'
        if ($Estatica) {
            $E = @($Estatica)[0]
            try { if ($E.DesignedCapacity -gt 0) { $Resultado.Diseno = [string][int]$E.DesignedCapacity } } catch {}
        }
    }

    if (-not $Resultado.Ciclos) {
        $Ciclos = Get-CimSeguro -Clase 'BatteryCycleCount' -Namespace 'root/wmi'
        if ($Ciclos) { try { $Resultado.Ciclos = [string]@($Ciclos)[0].CycleCount } catch {} }
    }

    $Estado = Get-CimSeguro -Clase 'BatteryStatus' -Namespace 'root/wmi'
    if ($Estado) {
        $S = @($Estado)[0]
        $Volt = 0; $Carga = 0; $Descarga = 0
        try { $Volt = [double]$S.Voltage } catch {}
        try { $Carga = [double]$S.ChargeRate } catch {}
        try { $Descarga = [double]$S.DischargeRate } catch {}

        if ($Volt -gt 0) { $Resultado.Voltaje = [string]([math]::Round($Volt / 1000, 2)) }
        $Potencia = if ($Carga -gt 0) { $Carga } else { $Descarga }
        if ($Potencia -gt 0) {
            $Resultado.Watts = [string]([math]::Round($Potencia / 1000, 1))
            if ($Volt -gt 0) { $Resultado.Amperaje = [string]([math]::Round(($Potencia / $Volt), 2)) }
        }
    }

    # Desgaste = perdida de capacidad respecto al diseno
    try {
        $D = [double]$Resultado.Diseno
        $A = [double]$Resultado.Actual
        if ($D -gt 0 -and $A -gt 0) {
            $Resultado.Desgaste = [string]([math]::Round((1 - ($A / $D)) * 100, 1))
        }
    } catch {}

    return [PSCustomObject]$Resultado
}

# -----------------------------------------------------------------------------
#  LICENCIAS
# -----------------------------------------------------------------------------
function Get-EstadoLicencia {
    $Resultado = [ordered]@{
        Windows = 'Desconocido'
        WinEdicion = ''
        Office = 'No instalado'
        OfficeVersion = ''
    }

    $AppIdWindows = '55c92734-d682-4d71-983e-d6ec3f16059f'
    $AppIdOffice = '0ff1ce15-a989-479d-af46-f275c6370663'

    # Consulta FILTRADA: enumerar todos los productos tarda ~18 s, filtrar tarda <1 s
    try {
        $Lic = Get-CimInstance -ClassName SoftwareLicensingProduct `
            -Filter "PartialProductKey IS NOT NULL" -ErrorAction Stop

        foreach ($L in $Lic) {
            $EstadoLic = switch ([int]$L.LicenseStatus) {
                0 { 'Sin licencia' }
                1 { 'Activado' }
                2 { 'Periodo de gracia' }
                3 { 'Gracia extendida' }
                4 { 'No genuino' }
                5 { 'Notificacion' }
                6 { 'Gracia extendida' }
                default { "Estado $($L.LicenseStatus)" }
            }

            if ("$($L.ApplicationId)" -eq $AppIdWindows) {
                if ($L.Name -notmatch 'Windows|Server') { continue }
                if ($Resultado.Windows -eq 'Desconocido' -or $EstadoLic -eq 'Activado') {
                    $Resultado.Windows = $EstadoLic
                    $Resultado.WinEdicion = ConvertTo-Plano $L.Name 60
                }
            }
            elseif ("$($L.ApplicationId)" -eq $AppIdOffice) {
                if ($Resultado.Office -eq 'No instalado' -or $EstadoLic -eq 'Activado') {
                    $Resultado.Office = $EstadoLic
                    $Resultado.OfficeVersion = ConvertTo-Plano $L.Name 60
                }
            }
        }
    } catch {}

    return [PSCustomObject]$Resultado
}

# -----------------------------------------------------------------------------
#  RED
# -----------------------------------------------------------------------------
function Get-RedDetalle {
    $Resultado = [ordered]@{
        Dominio    = 'N/A'
        IP         = 'N/A'
        MAC        = 'N/A'
        Gateway    = 'N/A'
        Adaptadores = 'N/A'
    }

    $CS = Get-CimSeguro -Clase 'Win32_ComputerSystem'
    if ($CS) {
        $Resultado.Dominio = if ($CS.PartOfDomain) { "Dominio: $($CS.Domain)" } else { "Grupo: $($CS.Workgroup)" }
    }

    $Adaptadores = Get-CimSeguro -Clase 'Win32_NetworkAdapterConfiguration' -Filtro 'IPEnabled=True'
    if ($Adaptadores) {
        $Lista = @($Adaptadores)
        $Principal = $null
        foreach ($a in $Lista) {
            if ($a.DefaultIPGateway) { $Principal = $a; break }
        }
        if (-not $Principal) { $Principal = $Lista[0] }

        $IPv4 = @($Principal.IPAddress) | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' } | Select-Object -First 1
        if ($IPv4) { $Resultado.IP = $IPv4 }
        if ($Principal.MACAddress) { $Resultado.MAC = $Principal.MACAddress }
        if ($Principal.DefaultIPGateway) { $Resultado.Gateway = (@($Principal.DefaultIPGateway) -join ', ') }

        $Nombres = $Lista | ForEach-Object {
            $ip = @($_.IPAddress) | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' } | Select-Object -First 1
            if ($ip) { '{0} ({1})' -f $_.Description, $ip }
        }
        $Resultado.Adaptadores = ConvertTo-Plano ($Nombres -join ' | ') 800
    }

    return [PSCustomObject]$Resultado
}

# -----------------------------------------------------------------------------
#  SEGURIDAD
# -----------------------------------------------------------------------------
function Get-SeguridadDetalle {
    $Resultado = [ordered]@{
        TPM        = 'No detectado'
        SecureBoot = 'No soportado'
        Antivirus  = 'N/A'
    }

    try {
        # Get-Tpm es ~50x mas rapido que la consulta CIM al namespace de TPM
        $Tpm2 = Get-Tpm -ErrorAction Stop
        if ($Tpm2.TpmPresent) {
            $Estado = if ($Tpm2.TpmEnabled) { 'Habilitado' } else { 'Deshabilitado' }
            $Version = $null
            foreach ($p in @('SpecVersion', 'TpmVersion')) {
                try { if ($Tpm2.$p) { $Version = [string]$Tpm2.$p; break } } catch {}
            }
            $Resultado.TPM = if ($Version) { '{0} v{1}' -f $Estado, $Version } else { $Estado }
        } else {
            $Resultado.TPM = 'No presente'
        }
    } catch {
        try {
            $Tpm = Get-CimInstance -Namespace 'root/cimv2/security/microsofttpm' -ClassName 'Win32_Tpm' -ErrorAction Stop
            if ($Tpm) {
                $T = @($Tpm)[0]
                $Estado = if ($T.IsEnabled_InitialValue) { 'Habilitado' } else { 'Deshabilitado' }
                $Resultado.TPM = ('{0} v{1}' -f $Estado, $T.SpecVersion).Trim()
            }
        } catch {}
    }
    if ($Resultado.TPM -eq 'No detectado' -and -not $script:EsAdmin) { $Resultado.TPM = 'Requiere admin' }

    try {
        $SB = Confirm-SecureBootUEFI -ErrorAction Stop
        $Resultado.SecureBoot = if ($SB) { 'Activado' } else { 'Desactivado' }
    } catch {
        $Resultado.SecureBoot = if ($script:EsAdmin) { 'No soportado' } else { 'Requiere admin' }
    }

    try {
        $AV = Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName 'AntiVirusProduct' -ErrorAction Stop
        if ($AV) {
            $Nombres = @($AV) | ForEach-Object {
                $activo = if (([int]$_.productState -band 0x1000) -ne 0) { 'activo' } else { 'inactivo' }
                '{0} ({1})' -f $_.displayName, $activo
            }
            $Resultado.Antivirus = ConvertTo-Plano ($Nombres -join ' | ') 300
        }
    } catch {}

    return [PSCustomObject]$Resultado
}

# -----------------------------------------------------------------------------
#  SOFTWARE INSTALADO
# -----------------------------------------------------------------------------
function Get-SoftwareInstalado {
    $Rutas = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $Items = @()
    foreach ($Ruta in $Rutas) {
        try {
            $Items += Get-ItemProperty -Path $Ruta -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.DisplayName -and
                    $_.SystemComponent -ne 1 -and
                    -not $_.ParentKeyName -and
                    -not $_.ReleaseType
                } |
                Select-Object DisplayName, DisplayVersion
        } catch {}
    }

    $Unicos = $Items |
        Group-Object DisplayName |
        ForEach-Object {
            $v = ($_.Group | Where-Object { $_.DisplayVersion } | Select-Object -First 1).DisplayVersion
            [PSCustomObject]@{
                Nombre   = $_.Name
                Version  = if ($v) { $v } else { '' }
            }
        } |
        Sort-Object Nombre

    return $Unicos
}

# -----------------------------------------------------------------------------
#  ESCANEO PRINCIPAL
# -----------------------------------------------------------------------------
function Get-InventoryRecord {
    param(
        [bool]$IncludeSoftware = $true
    )

    $R = [ordered]@{}

    foreach ($c in $script:Columnas) { $R[$c] = '' }

    $R['Fecha'] = (Get-Date -Format 'yyyy-MM-dd')
    $R['Hora']  = (Get-Date -Format 'HH:mm')

    # --- Identificacion ---
    $CS   = Get-CimSeguro -Clase 'Win32_ComputerSystem'
    $BIOS = Get-CimSeguro -Clase 'Win32_BIOS'
    $OS   = Get-CimSeguro -Clase 'Win32_OperatingSystem'
    $CPU  = Get-CimSeguro -Clase 'Win32_Processor' | Select-Object -First 1
    $Base = Get-CimSeguro -Clase 'Win32_BaseBoard'
    $Enc  = Get-CimSeguro -Clase 'Win32_SystemEnclosure' | Select-Object -First 1

    $R['Hostname'] = $env:COMPUTERNAME
    try { $R['Usuario'] = "$($CS.UserName)" } catch { $R['Usuario'] = $env:USERNAME }
    if (-not $R['Usuario']) { $R['Usuario'] = $env:USERNAME }

    $TiposChasis = @{
        3 = 'Desktop'; 4 = 'Desktop'; 5 = 'Desktop'; 6 = 'Desktop'; 7 = 'Desktop'
        8 = 'Portatil'; 9 = 'Laptop'; 10 = 'Notebook'; 14 = 'SubNotebook'
        13 = 'All-in-One'; 30 = 'Tablet'; 31 = 'Convertible'; 32 = 'Convertible'
    }
    $Tipo = 'PC'
    try {
        $Cod = [int]@($Enc.ChassisTypes)[0]
        if ($TiposChasis.ContainsKey($Cod)) { $Tipo = $TiposChasis[$Cod] }
    } catch {}
    $R['Tipo_Equipo'] = $Tipo

    $Marca = ConvertTo-Plano $CS.Manufacturer 80
    if ([string]::IsNullOrWhiteSpace($Marca) -or $Marca -match 'System manufacturer|OEM') {
        $Marca = ConvertTo-Plano $Base.Manufacturer 80
    }
    $R['Marca']  = $Marca
    $R['Modelo'] = ConvertTo-Plano $CS.Model 120

    $Serial = ConvertTo-Plano $BIOS.SerialNumber 80
    if ($Serial -match 'Default string|To be filled|O\.E\.M|None|0$') { $Serial = ConvertTo-Plano @($Enc.SerialNumber)[0] 80 }
    $R['Serial'] = $Serial
    $R['Serial_Chasis'] = ConvertTo-Plano @($Enc.SerialNumber)[0] 80
    $R['Asset_Tag'] = ConvertTo-Plano @($Enc.SMBIOSAssetTag)[0] 80

    # --- CPU ---
    if ($CPU) {
        $R['CPU_Modelo'] = ConvertTo-Plano ($CPU.Name -replace '\s+', ' ') 120
        $R['CPU_Nucleos'] = [string]$CPU.NumberOfCores
        $R['CPU_Hilos'] = [string]$CPU.NumberOfLogicalProcessors
        # MaxClockSpeed en equipos modernos reporta la frecuencia actual (throttling),
        # asi que se prefiere la frecuencia nominal impresa en el nombre del procesador.
        $GHz = $null
        if ($CPU.Name -match '@\s*([\d\.,]+)\s*GHz') {
            $Texto = $Matches[1].Replace(',', '.')
            $GHz = [double]::Parse($Texto, [System.Globalization.CultureInfo]::InvariantCulture)
        } elseif ($CPU.MaxClockSpeed -gt 0) {
            $GHz = [math]::Round($CPU.MaxClockSpeed / 1000, 2)
        }
        if ($GHz) { $R['CPU_GHz'] = [string]$GHz }
    }

    # --- RAM ---
    $RAM = Get-RAMDetalle
    $R['RAM_Total_GB'] = [string]$RAM.TotalGB
    $R['RAM_Tipo'] = $RAM.Tipo
    $R['RAM_Velocidad_MHz'] = $RAM.Velocidad
    $R['RAM_Slots'] = if ($RAM.SlotsTotales -gt 0) { '{0}/{1}' -f $RAM.SlotsUsados, $RAM.SlotsTotales } else { '' }
    $R['RAM_Slots_Libres'] = [string]$RAM.SlotsLibres
    $R['RAM_Max_GB'] = if ($RAM.MaxGB -gt 0) { [string]$RAM.MaxGB } else { '' }
    $R['RAM_Modulos'] = $RAM.Modulos

    # --- GPU ---
    $GPUs = Get-CimSeguro -Clase 'Win32_VideoController'
    if ($GPUs) {
        $Nombres = @($GPUs) | Where-Object { $_.Name } | ForEach-Object { ConvertTo-Plano $_.Name 80 }
        $R['GPU'] = ConvertTo-Plano ($Nombres -join ' | ') 300

        $VRAM = @($GPUs) | Where-Object { $_.AdapterRAM -gt 0 } | Select-Object -First 1
        if ($VRAM) { $R['GPU_VRAM'] = ('{0} GB' -f (ConvertTo-GB $VRAM.AdapterRAM 1)) }

        $Res = @($GPUs) | Where-Object { $_.CurrentHorizontalResolution -gt 0 } | Select-Object -First 1
        if ($Res) { $R['Resolucion'] = ('{0}x{1}' -f $Res.CurrentHorizontalResolution, $Res.CurrentVerticalResolution) }
    }

    $R['Monitores'] = ''
    try {
        if (-not ('System.Windows.Forms.Screen' -as [type])) { Add-Type -AssemblyName System.Windows.Forms }
        $Pantallas = [System.Windows.Forms.Screen]::AllScreens
        if ($Pantallas) { $R['Monitores'] = [string]@($Pantallas).Count }
    } catch {}

    # --- Placa / BIOS ---
    if ($Base) { $R['Placa_Madre'] = ConvertTo-Plano ('{0} {1}' -f $Base.Manufacturer, $Base.Product) 120 }
    if ($BIOS) {
        $R['BIOS_Version'] = ConvertTo-Plano $BIOS.SMBIOSBIOSVersion 80
        try { $R['BIOS_Fecha'] = ([datetime]$BIOS.ReleaseDate).ToString('yyyy-MM-dd') } catch {}
    }

    # --- Disco ---
    $Disco = Get-DiscoDetalle
    $R['Disco_Modelo'] = $Disco.Modelo
    $R['Disco_Tipo'] = $Disco.Tipo
    $R['Disco_Capacidad_GB'] = [string]$Disco.CapacidadGB
    $R['Disco_Libre_GB'] = [string]$Disco.LibreGB
    $R['Disco_Uso_Pct'] = $Disco.UsoPct
    $R['Disco_Salud'] = $Disco.Salud
    $R['Disco_Temperatura_C'] = $Disco.Temperatura
    $R['Disco_Horas_Encendido'] = $Disco.Horas
    $R['Discos_Extra'] = $Disco.Extra

    # --- Bateria ---
    $Bat = Get-BateriaDetalle
    $R['Bateria_Presente'] = $Bat.Presente
    $R['Bateria_Diseno_mWh'] = $Bat.Diseno
    $R['Bateria_Actual_mWh'] = $Bat.Actual
    $R['Bateria_Desgaste_Pct'] = $Bat.Desgaste
    $R['Bateria_Ciclos'] = $Bat.Ciclos
    $R['Bateria_Estado'] = $Bat.Estado
    $R['Bateria_Voltaje_V'] = $Bat.Voltaje
    $R['Bateria_Corriente_A'] = $Bat.Amperaje
    $R['Bateria_Potencia_W'] = $Bat.Watts

    # --- Sistema operativo ---
    if ($OS) {
        $R['SO_Nombre'] = ConvertTo-Plano $OS.Caption 100
        $R['SO_Version'] = [string]$OS.Version
        $R['SO_Build'] = [string]$OS.BuildNumber
        $R['SO_Arquitectura'] = [string]$OS.OSArchitecture
        try { $R['SO_Instalado'] = ([datetime]$OS.InstallDate).ToString('yyyy-MM-dd') } catch {}
        try {
            $Arranque = [datetime]$OS.LastBootUpTime
            $R['Ultimo_Arranque'] = $Arranque.ToString('yyyy-MM-dd HH:mm')
            $R['Uptime_Horas'] = [string]([math]::Round(((Get-Date) - $Arranque).TotalHours, 1))
        } catch {}
    }

    $Lic = Get-EstadoLicencia
    $R['Windows_Activado'] = $Lic.Windows
    $R['Office_Activado'] = $Lic.Office
    $R['Office_Version'] = $Lic.OfficeVersion

    # --- Red ---
    $Red = Get-RedDetalle
    $R['Dominio_Grupo'] = $Red.Dominio
    $R['IP'] = $Red.IP
    $R['MAC'] = $Red.MAC
    $R['Gateway'] = $Red.Gateway
    $R['Adaptadores'] = $Red.Adaptadores

    # --- Seguridad ---
    $Seg = Get-SeguridadDetalle
    $R['TPM'] = $Seg.TPM
    $R['SecureBoot'] = $Seg.SecureBoot
    $R['Antivirus'] = $Seg.Antivirus

    # --- Software ---
    if ($IncludeSoftware) {
        $Apps = Get-SoftwareInstalado
        $R['Software_Total'] = [string]@($Apps).Count
        $Lista = @($Apps) | ForEach-Object {
            if ($_.Version) { '{0} {1}' -f $_.Nombre, $_.Version } else { $_.Nombre }
        }
        $R['Software_Lista'] = ConvertTo-Plano ($Lista -join '; ') 30000
    }

    return [PSCustomObject]$R
}

# -----------------------------------------------------------------------------
#  CSV LOCAL
# -----------------------------------------------------------------------------
function Get-CsvPath {
    param([string]$DataDir, [string]$Cliente)
    $Nombre = if ([string]::IsNullOrWhiteSpace($Cliente)) { 'Inventario' } else { $Cliente }
    $Nombre = ($Nombre -replace '[\\/:*?"<>|]', '_').Trim()
    if ($Nombre.Length -gt 60) { $Nombre = $Nombre.Substring(0, 60) }
    return Join-Path $DataDir ('Inventario_{0}_{1}.csv' -f $Nombre, (Get-Date -Format 'yyyyMMdd'))
}

function Save-RecordCsv {
    param(
        [Parameter(Mandatory = $true)]$Registro,
        [Parameter(Mandatory = $true)][string]$DataDir,
        [string]$Cliente
    )

    if (-not (Test-Path $DataDir)) { New-Item -Path $DataDir -ItemType Directory -Force | Out-Null }

    $Csv = Get-CsvPath -DataDir $DataDir -Cliente $Cliente

    # Verificar que el encabezado existente coincida con el esquema actual
    if (Test-Path $Csv) {
        try {
            $Primera = Get-Content -Path $Csv -TotalCount 1 -Encoding UTF8
            $Esperado = '"' + ($script:Columnas -join '","') + '"'
            if ($Primera -ne $Esperado) {
                $Respaldo = $Csv -replace '\.csv$', ('_bak_{0}.csv' -f (Get-Date -Format 'HHmmss'))
                Move-Item -Path $Csv -Destination $Respaldo -Force
            }
        } catch {}
    }

    $Existe = Test-Path $Csv
    $Registro | Export-Csv -Path $Csv -NoTypeInformation -Encoding UTF8 -Append:$Existe

    return $Csv
}

# -----------------------------------------------------------------------------
#  COLA DE SINCRONIZACION (offline-first)
# -----------------------------------------------------------------------------
function Save-PendingRecord {
    param(
        [Parameter(Mandatory = $true)]$Registro,
        [Parameter(Mandatory = $true)][string]$QueueDir,
        [string]$Hoja
    )

    if (-not (Test-Path $QueueDir)) { New-Item -Path $QueueDir -ItemType Directory -Force | Out-Null }

    $Id = [guid]::NewGuid().ToString()
    $Payload = [ordered]@{
        id       = $Id
        schema   = 1
        hoja     = $Hoja
        registro = $Registro
        enviado  = $false
    }

    $Archivo = Join-Path $QueueDir ('{0}_{1}.json' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), $Id.Substring(0, 8))
    $Payload | ConvertTo-Json -Depth 5 | Out-File -FilePath $Archivo -Encoding UTF8 -Force

    return $Archivo
}

function Test-RegistroEnviado {
    param(
        [Parameter(Mandatory = $true)][string]$WebhookUrl,
        [string]$Token,
        [Parameter(Mandatory = $true)][string]$Id
    )

    # Devuelve $true / $false si el Apps Script confirma el id.
    # Devuelve $null si la verificacion no esta disponible (implementacion antigua).
    try {
        $Sep = if ($WebhookUrl.Contains('?')) { '&' } else { '?' }
        $Url = '{0}{1}verificar={2}&token={3}' -f $WebhookUrl, $Sep,
            [Uri]::EscapeDataString($Id), [Uri]::EscapeDataString($Token)

        $Req = [System.Net.HttpWebRequest]::Create($Url)
        $Req.Method = 'GET'
        $Req.Timeout = 20000
        $Req.ReadWriteTimeout = 20000
        $Req.AllowAutoRedirect = $true
        $Req.UserAgent = 'WigglesMini/1.0'

        $Resp = $Req.GetResponse()
        $Lector = New-Object System.IO.StreamReader($Resp.GetResponseStream(), [Text.Encoding]::UTF8)
        $Texto = $Lector.ReadToEnd()
        $Lector.Close()
        $Resp.Close()

        $Obj = $Texto | ConvertFrom-Json
        if ($Obj.PSObject.Properties.Name -contains 'ok' -and -not $Obj.ok) {
            return $false
        }
        if ($Obj.PSObject.Properties.Name -contains 'encontrado') {
            return [bool]$Obj.encontrado
        }
        return $null
    } catch {
        return $null
    }
}

function Send-PendingRecord {
    param(
        [Parameter(Mandatory = $true)][string]$Archivo,
        [Parameter(Mandatory = $true)][string]$WebhookUrl,
        [string]$Token
    )

    if (-not (Test-Path $Archivo)) { return $false }

    try {
        $Contenido = Get-Content -Path $Archivo -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $false
    }

    # El body incluye el token del cliente (el Apps Script lo valida)
    $Obj = [ordered]@{
        token    = $Token
        id       = $Contenido.id
        schema   = 1
        hoja     = $Contenido.hoja
        registro = $Contenido.registro
    }
    $Body = $Obj | ConvertTo-Json -Depth 6 -Compress

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $Req = [System.Net.HttpWebRequest]::Create($WebhookUrl)
        $Req.Method = 'POST'
        $Req.ContentType = 'application/json; charset=utf-8'
        $Req.Timeout = 20000
        $Req.ReadWriteTimeout = 20000
        $Req.AllowAutoRedirect = $false
        $Req.UserAgent = 'WigglesMini/1.0'

        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
        $Req.ContentLength = $Bytes.Length
        $Stream = $Req.GetRequestStream()
        $Stream.Write($Bytes, 0, $Bytes.Length)
        $Stream.Close()

        $Resp = $Req.GetResponse()
        $Codigo = [int]$Resp.StatusCode
        $Resp.Close()

        # Apps Script responde 302 al ejecutar doPost.
        # Antes de borrar la cola se confirma que el id llego al libro.
        if ($Codigo -eq 200 -or $Codigo -eq 302 -or $Codigo -eq 201) {
            $Verificado = Test-RegistroEnviado -WebhookUrl $WebhookUrl -Token $Token -Id $Contenido.id
            if ($Verificado -eq $false) { return $false }
            try { [IO.File]::Delete($Archivo) } catch {}
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

function Sync-PendingQueue {
    param(
        [Parameter(Mandatory = $true)][string]$QueueDir,
        [Parameter(Mandatory = $true)][string]$WebhookUrl,
        [string]$Token
    )

    $Res = [ordered]@{ Enviados = 0; Fallidos = 0; Pendientes = 0 }

    if (-not (Test-Path $QueueDir)) { return [PSCustomObject]$Res }

    $Archivos = Get-ChildItem -Path $QueueDir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($a in $Archivos) {
        if (Send-PendingRecord -Archivo $a.FullName -WebhookUrl $WebhookUrl -Token $Token) {
            $Res.Enviados++
        } else {
            $Res.Fallidos++
        }
    }

    $Res.Pendientes = @(Get-ChildItem -Path $QueueDir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    return [PSCustomObject]$Res
}

# -----------------------------------------------------------------------------
#  EXPORTAR A XLSX (OOXML nativo, sin modulos)
# -----------------------------------------------------------------------------
function ConvertTo-ColumnaExcel {
    param([int]$Indice)
    $s = ''
    while ($Indice -gt 0) {
        $r = ($Indice - 1) % 26
        $s = [char](65 + $r) + $s
        $Indice = [math]::Floor(($Indice - 1 - $r) / 26)
    }
    return $s
}

function ConvertTo-XmlSeguro {
    param([string]$Texto)
    if ($null -eq $Texto) { return '' }
    return [System.Security.SecurityElement]::Escape([string]$Texto)
}

function Export-CsvToXlsx {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$XlsxPath
    )

    if (-not (Test-Path $CsvPath)) { throw "No existe el CSV: $CsvPath" }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $Filas = @(Import-Csv -Path $CsvPath -Encoding UTF8)
    if ($Filas.Count -eq 0) { throw "El CSV esta vacio." }

    $Columnas = @($Filas[0].PSObject.Properties.Name)
    $TotalCols = $Columnas.Count
    $LetraFinal = ConvertTo-ColumnaExcel $TotalCols

    # --- Hoja ---
    $SB = New-Object System.Text.StringBuilder
    [void]$SB.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$SB.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$SB.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$SB.Append('<cols>')
    for ($c = 1; $c -le $TotalCols; $c++) {
        [void]$SB.Append(('<col min="{0}" max="{0}" width="18" customWidth="1"/>' -f $c))
    }
    [void]$SB.Append('</cols>')
    [void]$SB.Append('<sheetData>')

    # Encabezado
    [void]$SB.Append('<row r="1">')
    for ($c = 1; $c -le $TotalCols; $c++) {
        $Ref = (ConvertTo-ColumnaExcel $c) + '1'
        $Val = ConvertTo-XmlSeguro $Columnas[$c - 1]
        [void]$SB.Append(('<c r="{0}" t="inlineStr" s="1"><is><t xml:space="preserve">{1}</t></is></c>' -f $Ref, $Val))
    }
    [void]$SB.Append('</row>')

    # Datos
    $Fila = 1
    foreach ($Reg in $Filas) {
        $Fila++
        [void]$SB.Append(('<row r="{0}">' -f $Fila))
        for ($c = 1; $c -le $TotalCols; $c++) {
            $Ref = (ConvertTo-ColumnaExcel $c) + $Fila
            $Val = ConvertTo-XmlSeguro $Reg.($Columnas[$c - 1])
            if ($Val -ne '') {
                [void]$SB.Append(('<c r="{0}" t="inlineStr"><is><t xml:space="preserve">{1}</t></is></c>' -f $Ref, $Val))
            }
        }
        [void]$SB.Append('</row>')
    }

    [void]$SB.Append('</sheetData>')
    [void]$SB.Append(('<autoFilter ref="A1:{0}{1}"/>' -f $LetraFinal, $Fila))
    [void]$SB.Append('</worksheet>')

    $SheetXml = $SB.ToString()

    $ContentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>
'@

    $Rels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@

    $Workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="INVENTARIO" sheetId="1" r:id="rId1"/></sheets></workbook>
'@

    $WorkbookRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
'@

    $Styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F6F43"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>
'@

    if (Test-Path $XlsxPath) { try { [IO.File]::Delete($XlsxPath) } catch {} }

    $Stream = [System.IO.File]::Open($XlsxPath, [System.IO.FileMode]::Create)
    $Zip = New-Object System.IO.Compression.ZipArchive($Stream, [System.IO.Compression.ZipArchiveMode]::Create)

    $Entradas = @{
        '[Content_Types].xml'          = $ContentTypes
        '_rels/.rels'                  = $Rels
        'xl/workbook.xml'              = $Workbook
        'xl/_rels/workbook.xml.rels'   = $WorkbookRels
        'xl/styles.xml'                = $Styles
        'xl/worksheets/sheet1.xml'     = $SheetXml
    }

    $Utf8 = New-Object System.Text.UTF8Encoding($false)
    foreach ($Nombre in $Entradas.Keys) {
        $Entry = $Zip.CreateEntry($Nombre, [System.IO.Compression.CompressionLevel]::Optimal)
        $ES = $Entry.Open()
        $Bytes = $Utf8.GetBytes($Entradas[$Nombre])
        $ES.Write($Bytes, 0, $Bytes.Length)
        $ES.Close()
    }

    $Zip.Dispose()
    $Stream.Close()

    return $XlsxPath
}
