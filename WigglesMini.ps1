# =============================================================================
#  WIGGLES MINI  v1.0
#  Inventario de equipos + utilidades para visitas a clientes.
#  Un solo archivo, sin dependencias, 100% portable.
# =============================================================================
#Requires -Version 5.1

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$script:AppNombre  = 'WIGGLES MINI'
$script:AppVersion = '1.0.0'

# --- Ruta raiz (funciona desde USB, carpeta local o red) ---
$script:Root = $PSScriptRoot
if (-not $script:Root) {
    try { $script:Root = Split-Path -Parent $MyInvocation.MyCommand.Definition } catch {}
}
if (-not $script:Root) { $script:Root = (Get-Location).Path }
$script:Root = $script:Root.TrimEnd('\')

$script:CorePath   = Join-Path $script:Root 'WigglesCore.ps1'
$script:DataDir    = Join-Path $script:Root 'Data'
$script:LogDir     = Join-Path $script:DataDir 'Logs'
$script:QueueDir   = Join-Path $script:DataDir 'Pendientes'
$script:ToolsDir   = Join-Path $script:Root 'External\Tools'
$script:OfficeDir  = Join-Path $script:Root 'External\Office'

foreach ($d in @($script:DataDir, $script:LogDir, $script:QueueDir)) {
    if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
}

# --- Elevacion ---
function Test-EsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

if (-not (Test-EsAdmin)) {
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`"" -Verb RunAs
    } catch {}
    exit
}

# --- Log ---
$script:LogFile = Join-Path $script:LogDir ('Mini_{0}.log' -f (Get-Date -Format 'yyyy-MM-dd'))
try {
    Get-ChildItem -Path $script:LogDir -Filter 'Mini_*.log' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 15 |
        ForEach-Object { try { [IO.File]::Delete($_.FullName) } catch {} }
} catch {}

function Write-Log {
    param([string]$Mensaje, [string]$Nivel = 'INFO')
    $Linea = '{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Nivel, $Mensaje
    try { Add-Content -Path $script:LogFile -Value $Linea -Encoding UTF8 } catch {}
    if ($script:txtLog) {
        $script:txtLog.AppendText($Linea + "`r`n")
        $script:txtLog.ScrollToEnd()
    }
}

# --- Núcleo ---
if (-not (Test-Path $script:CorePath)) {
    [System.Windows.Forms.MessageBox]::Show("No se encontro WigglesCore.ps1 junto a este archivo.`n`nCarpeta: $script:Root", 'Wiggles Mini', 'OK', 'Error')
    exit 1
}
. $script:CorePath

# --- Configuracion ---
$script:ConfigPath = Join-Path $script:Root 'config.json'
$script:Config = [ordered]@{
    Cliente             = ''
    Tecnico             = $env:USERNAME
    WebhookUrl          = ''
    Token               = ''
    AppsBasicas         = @('7zip.7zip', 'Google.Chrome', 'VideoLAN.VLC')
    SincronizarAlGuardar = $true
}

if (Test-Path $script:ConfigPath) {
    try {
        $Leido = Get-Content -Path $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in @($script:Config.Keys)) {
            if ($null -ne $Leido.$k) { $script:Config[$k] = $Leido.$k }
        }
    } catch {
        Write-Log "config.json invalido: $($_.Exception.Message)" 'WARN'
    }
}

function Save-Config {
    try {
        $script:Config | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:ConfigPath -Encoding UTF8 -Force
    } catch { Write-Log "No se pudo guardar config.json" 'WARN' }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============================================================================
#  INTERFAZ
# =============================================================================
$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Wiggles Mini" Height="800" Width="1220" MinHeight="660" MinWidth="1020"
        WindowStartupLocation="CenterScreen" WindowStyle="None"
        Background="#0B0F14" Foreground="#E6EDF3"
        FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True"
        TextOptions.TextFormattingMode="Display">

  <shell:WindowChrome.WindowChrome>
    <shell:WindowChrome CaptionHeight="46" ResizeBorderThickness="6" CornerRadius="0" GlassFrameThickness="0"/>
  </shell:WindowChrome.WindowChrome>

  <Window.Resources>
    <SolidColorBrush x:Key="Fondo" Color="#0B0F14"/>
    <SolidColorBrush x:Key="Panel" Color="#121821"/>
    <SolidColorBrush x:Key="PanelAlto" Color="#18202B"/>
    <SolidColorBrush x:Key="Borde" Color="#1E2A38"/>
    <SolidColorBrush x:Key="Verde" Color="#00E676"/>
    <SolidColorBrush x:Key="Azul" Color="#29B6F6"/>
    <SolidColorBrush x:Key="Texto" Color="#E6EDF3"/>
    <SolidColorBrush x:Key="Tenue" Color="#8494A6"/>
    <SolidColorBrush x:Key="Ambar" Color="#FFB300"/>
    <SolidColorBrush x:Key="Rojo" Color="#FF5252"/>

    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Texto}"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>

    <Style x:Key="Etiqueta" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Tenue}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Margin" Value="0,0,0,3"/>
    </Style>

    <Style x:Key="Titulo" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Verde}"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#0F1520"/>
      <Setter Property="Foreground" Value="{StaticResource Texto}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Borde}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="CaretBrush" Value="{StaticResource Verde}"/>
      <Setter Property="SelectionBrush" Value="#2E7D32"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Verde}"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#2C3B4D"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="Button">
      <Setter Property="Background" Value="#18202B"/>
      <Setter Property="Foreground" Value="{StaticResource Texto}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Borde}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="14,9"/>
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#1F2B3A"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="#2E4053"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#0D141C"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BotonVerde" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#0E3B24"/>
      <Setter Property="BorderBrush" Value="#1B7A46"/>
      <Setter Property="Foreground" Value="#B9F6CA"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <Style x:Key="BotonAzul" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#0E2A3B"/>
      <Setter Property="BorderBrush" Value="#22688F"/>
      <Setter Property="Foreground" Value="#B3E5FC"/>
    </Style>

    <Style x:Key="BotonAmbar" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#3B2E0E"/>
      <Setter Property="BorderBrush" Value="#8F6B22"/>
      <Setter Property="Foreground" Value="#FFE082"/>
    </Style>

    <Style x:Key="BotonFantasma" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10,6"/>
    </Style>

    <Style x:Key="Tarjeta" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource Panel}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Borde}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="16"/>
      <Setter Property="Margin" Value="0,0,12,12"/>
    </Style>

    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="{StaticResource Tenue}"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" Padding="18,10" Margin="0,0,4,0"
                    CornerRadius="6,6,0,0" BorderThickness="0,0,0,2" BorderBrush="Transparent">
              <ContentPresenter ContentSource="Header" HorizontalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Verde}"/>
                <Setter Property="Foreground" Value="{StaticResource Verde}"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="{StaticResource Texto}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="{StaticResource Borde}" BorderThickness="0,0,0,1">
                <TabPanel IsItemsHost="True" Margin="14,0,0,0"/>
              </Border>
              <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="#0F1520"/>
      <Setter Property="Foreground" Value="{StaticResource Texto}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Borde}"/>
      <Setter Property="Padding" Value="8,6"/>
    </Style>

    <Style x:Key="BarraProgreso" TargetType="ProgressBar">
      <Setter Property="Height" Value="3"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource Verde}"/>
      <Setter Property="BorderThickness" Value="0"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="46"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- ================= ENCABEZADO ================= -->
    <Border Grid.Row="0" Background="#0E141C" BorderBrush="{StaticResource Borde}" BorderThickness="0,0,0,1">
      <Grid Margin="16,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="8" Height="22" Background="{StaticResource Verde}" CornerRadius="2" Margin="0,0,10,0"/>
          <TextBlock Text="WIGGLES" FontSize="16" FontWeight="Bold" Foreground="{StaticResource Texto}"/>
          <TextBlock Text="MINI" FontSize="16" FontWeight="Bold" Foreground="{StaticResource Verde}" Margin="6,0,0,0"/>
          <Border Background="#1B2634" CornerRadius="3" Padding="6,2" Margin="10,0,0,0" VerticalAlignment="Center">
            <TextBlock Name="txtVersion" Text="v1.0" FontSize="10" Foreground="{StaticResource Tenue}"/>
          </Border>
        </StackPanel>

        <TextBlock Grid.Column="1" Name="txtEncabezado" Text="Inventario para visitas a clientes"
                   Foreground="{StaticResource Tenue}" FontSize="11.5" Margin="24,0,0,0"/>

        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center"
                    shell:WindowChrome.IsHitTestVisibleInChrome="True">
          <Button Name="btnMin" Content="—" Style="{StaticResource BotonFantasma}" FontSize="14"
                  Width="42" Height="30" Margin="0" ToolTip="Minimizar"/>
          <Button Name="btnCerrar" Content="✕" Style="{StaticResource BotonFantasma}" FontSize="13"
                  Width="42" Height="30" Margin="0" ToolTip="Cerrar"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ================= CONTENIDO ================= -->
    <TabControl Grid.Row="1" Name="tabs" Margin="6,4,6,0">

      <!-- ===== INVENTARIO ===== -->
      <TabItem Header="INVENTARIO">
        <Grid Margin="14,14,14,6">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="430"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Margin="0,0,14,0">
            <StackPanel>
              <TextBlock Text="DATOS DE LA VISITA" Style="{StaticResource Titulo}"/>

              <TextBlock Text="CLIENTE" Style="{StaticResource Etiqueta}"/>
              <TextBox Name="txtCliente" Margin="0,0,0,9"/>

              <Grid Margin="0,0,0,9">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock Text="SEDE / TECNICO" Style="{StaticResource Etiqueta}"/>
                  <TextBox Name="txtSede"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="ORDEN / ETIQUETA" Style="{StaticResource Etiqueta}"/>
                  <TextBox Name="txtOrden"/>
                </StackPanel>
              </Grid>

              <Border Height="1" Background="{StaticResource Borde}" Margin="0,4,0,14"/>

              <TextBlock Text="EQUIPO" Style="{StaticResource Titulo}"/>

              <TextBlock Text="SERIAL / SERVICE TAG" Style="{StaticResource Etiqueta}"/>
              <TextBox Name="txtSerial" FontSize="15" FontWeight="Bold" Foreground="{StaticResource Verde}"
                       Margin="0,0,0,9"/>

              <Grid Margin="0,0,0,9">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock Text="MARCA" Style="{StaticResource Etiqueta}"/>
                  <TextBox Name="txtMarca"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="MODELO" Style="{StaticResource Etiqueta}"/>
                  <TextBox Name="txtModelo"/>
                </StackPanel>
              </Grid>

              <TextBlock Text="COMENTARIOS / ESTADO" Style="{StaticResource Etiqueta}"/>
              <TextBox Name="txtComentario" Height="66" AcceptsReturn="True" TextWrapping="Wrap"
                       VerticalScrollBarVisibility="Auto" Margin="0,0,0,14"/>

              <Button Name="btnEscanear" Content="ESCANEAR EQUIPO" Style="{StaticResource BotonAzul}"
                      Height="38" Margin="0,0,0,8"/>
              <Button Name="btnGuardarSiguiente" Content="GUARDAR Y SIGUIENTE EQUIPO ▸"
                      Style="{StaticResource BotonVerde}" Height="42" FontSize="13" Margin="0,0,0,8"/>
              <Button Name="btnGuardar" Content="GUARDAR SOLO ESTE EQUIPO" Height="34" Margin="0,0,0,8"/>

              <WrapPanel>
                <Button Name="btnExcel" Content="Exportar Excel" Height="30" FontSize="11.5"/>
                <Button Name="btnSincronizar" Content="Sincronizar" Height="30" FontSize="11.5"/>
                <Button Name="btnCarpeta" Content="Abrir carpeta" Height="30" FontSize="11.5"/>
              </WrapPanel>

              <Border Background="#101822" BorderBrush="{StaticResource Borde}" BorderThickness="1"
                      CornerRadius="6" Padding="10,8" Margin="0,6,0,0">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="Sesion:" Foreground="{StaticResource Tenue}" FontSize="11.5"/>
                  <TextBlock Name="txtContador" Text="0 equipos" Foreground="{StaticResource Verde}"
                             FontSize="11.5" FontWeight="SemiBold" Margin="6,0,0,0"/>
                  <TextBlock Name="txtPendientes" Text="" Foreground="{StaticResource Ambar}"
                             FontSize="11.5" Margin="16,0,0,0"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </ScrollViewer>

          <Border Grid.Column="1" Style="{StaticResource Tarjeta}" Margin="0,0,0,10">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                <TextBlock Text="FICHA DEL EQUIPO" Style="{StaticResource Titulo}" Margin="0"/>
                <TextBlock Name="txtScanEstado" Text="sin escanear" Foreground="{StaticResource Tenue}"
                           FontSize="11" Margin="12,0,0,0"/>
              </StackPanel>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                <StackPanel Name="pnlHardware"/>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- ===== SISTEMA ===== -->
      <TabItem Header="SISTEMA">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <Grid Margin="14,14,14,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition/>
              <ColumnDefinition/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
              <TextBlock Text="LICENCIAMIENTO" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <TextBlock Name="txtLicencia" Text="Windows: detectando...   Office: detectando..."
                             Foreground="{StaticResource Tenue}" FontSize="11.5" Margin="0,0,0,12" TextWrapping="Wrap"/>
                  <Button Name="btnActivar" Content="ACTIVAR WINDOWS Y OFFICE" Style="{StaticResource BotonAmbar}"
                          Height="40" Margin="0,0,0,8"/>
                  <Button Name="btnEdicion" Content="CAMBIAR EDICION  (Home → Pro)" Height="34" Margin="0,0,0,4"/>
                  <TextBlock Text="Requiere reinicio al finalizar. La activacion necesita conexion a internet."
                             Foreground="{StaticResource Tenue}" FontSize="10.5" TextWrapping="Wrap" Margin="0,4,0,0"/>
                </StackPanel>
              </Border>

              <TextBlock Text="MANTENIMIENTO" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <Button Name="btnPuntoRestauracion" Content="Crear punto de restauracion" Margin="0,0,0,6"/>
                  <Button Name="btnBloatware" Content="Eliminar bloatware (apps preinstaladas)" Margin="0,0,0,6"/>
                  <Button Name="btnRendimiento" Content="Plan de maximo rendimiento" Margin="0,0,0,6"/>
                  <Button Name="btnRed" Content="Reparar red (winsock + DNS)" Margin="0,0,0,6"/>
                  <Button Name="btnAdminLocal" Content="Crear usuario administrador local" Margin="0,0,0,0"/>
                </StackPanel>
              </Border>
            </StackPanel>

            <StackPanel Grid.Column="1">
              <TextBlock Text="DIAGNOSTICO" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <Button Name="btnClaveBios" Content="Ver clave de Windows en BIOS (OA3)" Margin="0,0,0,6"/>
                  <Button Name="btnReporteBateria" Content="Generar reporte de bateria (HTML)" Margin="0,0,0,6"/>
                  <Button Name="btnResumenEquipo" Content="Ver resumen completo del equipo" Margin="0,0,0,0"/>
                </StackPanel>
              </Border>

              <TextBlock Text="EMPAQUETADO" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <Button Name="btnExportarTodo" Content="Exportar inventario del dia a Excel" Margin="0,0,0,6"/>
                  <Button Name="btnSyncTodo" Content="Sincronizar todo con Google Sheets" Margin="0,0,0,10"/>
                  <TextBlock Text="Los registros se guardan siempre en el pendrive. La nube es opcional y se reintenta sola cuando hay internet."
                             Foreground="{StaticResource Tenue}" FontSize="10.5" TextWrapping="Wrap"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- ===== SOFTWARE ===== -->
      <TabItem Header="SOFTWARE">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <Grid Margin="14,14,14,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
              <TextBlock Text="MICROSOFT OFFICE" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <TextBlock Name="txtOfficeEstado" Text="Verificando..." Foreground="{StaticResource Tenue}"
                             FontSize="11.5" TextWrapping="Wrap" Margin="0,0,0,10"/>
                  <Button Name="btnOffice" Content="INSTALAR OFFICE 2024" Style="{StaticResource BotonVerde}"
                          Height="40" Margin="0,0,0,8"/>
                  <TextBlock Text="Office 2024 se descarga de internet durante la instalacion (~2 GB). Necesita conexion estable."
                             Foreground="{StaticResource Tenue}" FontSize="10.5" TextWrapping="Wrap" Margin="0,4,0,0"/>
                </StackPanel>
              </Border>

              <TextBlock Text="APLICACIONES BASICAS" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <TextBlock Name="txtAppsLista" Text="7-Zip · Chrome · VLC" Foreground="{StaticResource Tenue}"
                             FontSize="11.5" TextWrapping="Wrap" Margin="0,0,0,10"/>
                  <Button Name="btnApps" Content="INSTALAR APLICACIONES BASICAS" Style="{StaticResource BotonAzul}"
                          Height="36" Margin="0,0,0,4"/>
                </StackPanel>
              </Border>
            </StackPanel>

            <StackPanel Grid.Column="1">
              <TextBlock Text="COMPLEMENTOS" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <Button Name="btnRuntimes" Content="Instalar runtimes (Visual C++ / .NET)" Margin="0,0,0,6"/>
                  <Button Name="btnQuitarOffice" Content="Desinstalar Office existente" Margin="0,0,0,6"/>
                  <Button Name="btnWingetReparar" Content="Reparar winget / App Installer" Margin="0,0,0,0"/>
                </StackPanel>
              </Border>

              <TextBlock Text="INFORMACION" Style="{StaticResource Titulo}"/>
              <Border Style="{StaticResource Tarjeta}">
                <StackPanel>
                  <TextBlock Text="Las instalaciones por winget y Office requieren conexion a internet."
                             Foreground="{StaticResource Tenue}" FontSize="11" TextWrapping="Wrap"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- ===== HERRAMIENTAS ===== -->
      <TabItem Header="HERRAMIENTAS">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="14,14,14,10">
            <TextBlock Text="HERRAMIENTAS PORTABLES" Style="{StaticResource Titulo}"/>
            <TextBlock Name="txtToolsInfo" Text="" Foreground="{StaticResource Tenue}" FontSize="11"
                       Margin="0,-4,0,12" TextWrapping="Wrap"/>
            <WrapPanel Name="pnlTools"/>
          </StackPanel>
        </ScrollViewer>
      </TabItem>
    </TabControl>

    <!-- ================= PIE ================= -->
    <Border Grid.Row="2" Background="#0E141C" BorderBrush="{StaticResource Borde}" BorderThickness="0,1,0,0">
      <Grid Margin="16,8,16,8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Orientation="Horizontal">
            <Border Name="bdEstado" Width="8" Height="8" CornerRadius="4" Background="{StaticResource Verde}"
                    VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock Name="txtEstado" Text="Listo" Foreground="{StaticResource Tenue}" FontSize="11.5"/>
          </StackPanel>
          <TextBlock Grid.Column="1" Name="txtRuta" Text="" Foreground="#4A5A6B" FontSize="10.5"/>
        </Grid>

        <ProgressBar Grid.Row="1" Name="prgBarra" Style="{StaticResource BarraProgreso}"
                     IsIndeterminate="False" Visibility="Hidden" Margin="0,6,0,0"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

try {
    $Window = [Windows.Markup.XamlReader]::Parse($Xaml)
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error al construir la interfaz:`n`n$($_.Exception.Message)", 'Wiggles Mini', 'OK', 'Error')
    exit 1
}

# Ajustar la ventana al area util de la pantalla (evita que se salga con escalado DPI)
try {
    $Area = [System.Windows.SystemParameters]::WorkArea
    $Ancho = [math]::Min(1220, [math]::Max(1020, $Area.Width - 24))
    $Alto  = [math]::Min(800,  [math]::Max(620, $Area.Height - 24))
    $Window.Width  = $Ancho
    $Window.Height = $Alto
    $Window.Left   = $Area.Left + [math]::Round(($Area.Width - $Ancho) / 2)
    $Window.Top    = $Area.Top  + [math]::Round(($Area.Height - $Alto) / 2)
} catch {}

# Conectar controles
try { [xml]$XamlDoc = $Xaml } catch {}
$XamlDoc.SelectNodes('//*[@Name]') | ForEach-Object {
    try {
        $ctrl = $Window.FindName($_.Name)
        if ($ctrl) { Set-Variable -Name $_.Name -Value $ctrl -Scope Script }
    } catch {}
}

$script:txtLog = $null
$txtVersion.Text = $script:AppVersion
$txtRuta.Text = $script:Root

# --- Barra de estado ---
function Set-Estado {
    param([string]$Texto, [string]$Tipo = 'ok')
    $txtEstado.Text = $Texto
    $Color = switch ($Tipo) {
        'ok'    { '#00E676' }
        'info'  { '#29B6F6' }
        'warn'  { '#FFB300' }
        'error' { '#FF5252' }
        default { '#00E676' }
    }
    try { $bdEstado.Background = $Color } catch {}
}

$script:TareasActivas = 0

function Set-Ocupado {
    param([bool]$Activo)
    if ($Activo) {
        $Botonera = @('btnEscanear', 'btnGuardar', 'btnGuardarSiguiente', 'btnExcel', 'btnSincronizar',
                      'btnOffice', 'btnApps', 'btnActivar', 'btnEdicion', 'btnExportarTodo', 'btnSyncTodo')
        foreach ($b in $Botonera) {
            try { $Window.FindName($b).IsEnabled = $false } catch {}
        }
        $prgBarra.IsIndeterminate = $true
        $prgBarra.Visibility = 'Visible'
    } else {
        $Botonera = @('btnEscanear', 'btnGuardar', 'btnGuardarSiguiente', 'btnExcel', 'btnSincronizar',
                      'btnOffice', 'btnApps', 'btnActivar', 'btnEdicion', 'btnExportarTodo', 'btnSyncTodo')
        foreach ($b in $Botonera) {
            try { $Window.FindName($b).IsEnabled = $true } catch {}
        }
        $prgBarra.IsIndeterminate = $false
        $prgBarra.Visibility = 'Hidden'
    }
}

# =============================================================================
#  TAREAS EN SEGUNDO PLANO
# =============================================================================
$script:Tareas = New-Object System.Collections.ArrayList

$script:TimerTareas = New-Object System.Windows.Threading.DispatcherTimer
$script:TimerTareas.Interval = [TimeSpan]::FromMilliseconds(250)
$script:TimerTareas.Add_Tick({
    if ($script:Tareas.Count -eq 0) { return }
    for ($i = $script:Tareas.Count - 1; $i -ge 0; $i--) {
        $t = $script:Tareas[$i]
        if ($t.Handle.IsCompleted) {
            $script:Tareas.RemoveAt($i)
            $resultado = $null
            $errorTxt = $null
            try {
                $resultado = $t.PS.EndInvoke($t.Handle)
            } catch {
                $errorTxt = $_.Exception.Message
                try {
                    $detalle = $t.PS.Streams.Error | ForEach-Object { $_.ToString() }
                    if ($detalle) { $errorTxt = $detalle -join ' | ' }
                } catch {}
            }
            try { $t.PS.Dispose() } catch {}
            try { $t.RS.Dispose() } catch {}
            $script:TareasActivas = [math]::Max(0, $script:TareasActivas - 1)
            if ($script:TareasActivas -eq 0) { Set-Ocupado $false }
            if ($t.OnDone) {
                try { & $t.OnDone $resultado $errorTxt } catch { Write-Log "Error en callback: $($_.Exception.Message)" 'ERROR' }
            }
        }
    }
})
$script:TimerTareas.Start()

function Start-Tarea {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Script,
        [object[]]$Argumentos = @(),
        [scriptblock]$OnDone
    )
    $ps = [powershell]::Create()
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Script)
    foreach ($a in $Argumentos) { [void]$ps.AddArgument($a) }
    $h = $ps.BeginInvoke()
    [void]$script:Tareas.Add([PSCustomObject]@{ PS = $ps; RS = $rs; Handle = $h; OnDone = $OnDone })
    $script:TareasActivas++
    Set-Ocupado $true
}

# =============================================================================
#  CAMPOS DE LA FICHA (panel derecho)
# =============================================================================
$script:CamposFicha = New-Object System.Collections.Specialized.OrderedDictionary

function New-FilaFicha {
    param([string]$Etiqueta, [string]$Clave, [bool]$Destacado = $false)
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = '0,0,0,7'
    $c1 = New-Object System.Windows.Controls.ColumnDefinition
    $c1.Width = New-Object System.Windows.GridLength(158)
    $c2 = New-Object System.Windows.Controls.ColumnDefinition
    $c2.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $grid.ColumnDefinitions.Add($c1)
    $grid.ColumnDefinitions.Add($c2)

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = $Etiqueta
    $lbl.FontSize = 11
    $lbl.Foreground = '#8494A6'
    $lbl.TextWrapping = 'Wrap'
    $lbl.VerticalAlignment = 'Top'
    $lbl.Margin = '0,4,8,0'
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
    $grid.Children.Add($lbl) | Out-Null

    $val = New-Object System.Windows.Controls.TextBlock
    $val.Text = ''
    $val.FontSize = 12
    $val.TextWrapping = 'Wrap'
    $val.Foreground = if ($Destacado) { '#E6EDF3' } else { '#B8C6D4' }
    if ($Destacado) { $val.FontWeight = 'SemiBold' }
    [System.Windows.Controls.Grid]::SetColumn($val, 1)
    $grid.Children.Add($val) | Out-Null

    $script:CamposFicha[$Clave] = $val
    return $grid
}

function Set-Ficha {
    param([string]$Clave, [string]$Valor)
    if ($script:CamposFicha.Contains($Clave)) {
        $script:CamposFicha[$Clave].Text = if ([string]::IsNullOrWhiteSpace($Valor)) { '—' } else { $Valor }
    }
}

$DefFicha = @(
    [PSCustomObject]@{ E = 'Fabricante y modelo'; C = 'fMarca';     D = $true }
    [PSCustomObject]@{ E = 'Tipo de equipo';      C = 'fTipo';      D = $false }
    [PSCustomObject]@{ E = 'Procesador';          C = 'fCPU';       D = $true }
    [PSCustomObject]@{ E = 'Nucleos / hilos';     C = 'fCores';     D = $false }
    [PSCustomObject]@{ E = 'Memoria RAM';         C = 'fRAM';       D = $true }
    [PSCustomObject]@{ E = 'Tipo / velocidad';    C = 'fRAMTipo';   D = $false }
    [PSCustomObject]@{ E = 'Modulos';             C = 'fRAMMod';    D = $false }
    [PSCustomObject]@{ E = 'Slots libres / max';  C = 'fSlots';     D = $false }
    [PSCustomObject]@{ E = 'Disco';               C = 'fDisco';     D = $true }
    [PSCustomObject]@{ E = 'Tipo / capacidad';    C = 'fDiscoTipo'; D = $false }
    [PSCustomObject]@{ E = 'Espacio libre';       C = 'fLibre';     D = $false }
    [PSCustomObject]@{ E = 'Salud / temperatura'; C = 'fSalud';     D = $false }
    [PSCustomObject]@{ E = 'Bateria';             C = 'fBateria';   D = $true }
    [PSCustomObject]@{ E = 'Desgaste / ciclos';   C = 'fDesgaste';  D = $false }
    [PSCustomObject]@{ E = 'Sistema operativo';   C = 'fSO';        D = $true }
    [PSCustomObject]@{ E = 'Version / build';     C = 'fBuild';     D = $false }
    [PSCustomObject]@{ E = 'Windows';             C = 'fWinLic';    D = $false }
    [PSCustomObject]@{ E = 'Office';              C = 'fOfficeLic'; D = $false }
    [PSCustomObject]@{ E = 'Dominio / grupo';     C = 'fDominio';   D = $false }
    [PSCustomObject]@{ E = 'IP / MAC';            C = 'fRed';       D = $false }
    [PSCustomObject]@{ E = 'TPM';                 C = 'fTPM';       D = $false }
    [PSCustomObject]@{ E = 'Secure Boot';         C = 'fSB';        D = $false }
    [PSCustomObject]@{ E = 'Software instalado';  C = 'fSoftware';  D = $false }
    [PSCustomObject]@{ E = 'Ultimo arranque';     C = 'fArranque';  D = $false }
)

foreach ($d in $DefFicha) {
    $pnlHardware.Children.Add((New-FilaFicha -Etiqueta $d.E -Clave $d.C -Destacado $d.D)) | Out-Null
}

function Reset-Ficha {
    foreach ($k in @($script:CamposFicha.Keys)) { $script:CamposFicha[$k].Text = '' }
}

function Mostrar-Ficha {
    param($Info)
    if (-not $Info) { return }
    Set-Ficha 'fMarca'  ("{0} {1}" -f $Info.Marca, $Info.Modelo).Trim()
    Set-Ficha 'fTipo'   $Info.Tipo_Equipo
    Set-Ficha 'fCPU'    $Info.CPU_Modelo
    Set-Ficha 'fCores'  ("{0} nucleos / {1} hilos · {2} GHz" -f $Info.CPU_Nucleos, $Info.CPU_Hilos, $Info.CPU_GHz).Trim(' ', '/', '·')
    Set-Ficha 'fRAM'    ("{0} GB" -f $Info.RAM_Total_GB)
    Set-Ficha 'fRAMTipo' ("{0} · {1}" -f $Info.RAM_Tipo, $Info.RAM_Velocidad_MHz).Trim(' ', '·')
    Set-Ficha 'fRAMMod' $Info.RAM_Modulos
    $SlotsTotal = 0
    if ("$($Info.RAM_Slots)" -match '/(\d+)') { $SlotsTotal = [int]$Matches[1] }
    $SlotsTxt = if ($SlotsTotal -gt 0) { "{0} libres de {1}" -f $Info.RAM_Slots_Libres, $SlotsTotal } else { "$($Info.RAM_Slots)" }
    if ([int]$Info.RAM_Max_GB -gt 0) { $SlotsTxt = ("{0} · max {1} GB" -f $SlotsTxt, $Info.RAM_Max_GB) }
    Set-Ficha 'fSlots' $SlotsTxt
    Set-Ficha 'fDisco'  $Info.Disco_Modelo
    Set-Ficha 'fDiscoTipo' ("{0} · {1} GB" -f $Info.Disco_Tipo, $Info.Disco_Capacidad_GB).Trim(' ', '·')
    Set-Ficha 'fLibre'  ("{0} GB libres ({1}% usado)" -f $Info.Disco_Libre_GB, $Info.Disco_Uso_Pct)
    Set-Ficha 'fSalud'  ("{0} {1}" -f $Info.Disco_Salud, $(if ($Info.Disco_Temperatura_C) { "· $($Info.Disco_Temperatura_C) C" } else { '' })).Trim()
    Set-Ficha 'fBateria' $Info.Bateria_Estado
    Set-Ficha 'fDesgaste' ("{0}% desgaste · {1} ciclos" -f $Info.Bateria_Desgaste_Pct, $Info.Bateria_Ciclos).Trim('%', ' ', '·')
    Set-Ficha 'fSO'     $Info.SO_Nombre
    Set-Ficha 'fBuild'  ("{0} (build {1})" -f $Info.SO_Version, $Info.SO_Build)
    Set-Ficha 'fWinLic' $Info.Windows_Activado
    Set-Ficha 'fOfficeLic' ("{0} {1}" -f $Info.Office_Activado, $Info.Office_Version).Trim()
    Set-Ficha 'fDominio' $Info.Dominio_Grupo
    Set-Ficha 'fRed'    ("{0} · {1}" -f $Info.IP, $Info.MAC).Trim(' ', '·')
    Set-Ficha 'fTPM'    $Info.TPM
    Set-Ficha 'fSB'     $Info.SecureBoot
    Set-Ficha 'fSoftware' ("{0} programas" -f $Info.Software_Total)
    Set-Ficha 'fArranque' $Info.Ultimo_Arranque
}

# =============================================================================
#  ESTADO DE LA SESION
# =============================================================================
$script:RegistroActual = $null
$script:ContadorSesion = 0

function Update-Contadores {
    $txtContador.Text = '{0} equipo{1}' -f $script:ContadorSesion, $(if ($script:ContadorSesion -eq 1) { '' } else { 's' })
    $pend = 0
    try { $pend = @(Get-ChildItem -Path $script:QueueDir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count } catch {}
    $txtPendientes.Text = if ($pend -gt 0) { "· $pend pendiente(s) de sincronizar" } else { '' }
}

function Get-HojaActual {
    $cliente = $txtCliente.Text.Trim()
    $nombre = if ([string]::IsNullOrWhiteSpace($cliente)) { 'Sin cliente' } else { $cliente }
    return ('{0} · {1}' -f (Get-Date -Format 'yyyy-MM-dd'), $nombre)
}

# =============================================================================
#  ESCANEO
# =============================================================================
function Start-Escaneo {
    if ($script:TareasActivas -gt 0) {
        Set-Estado 'Espera a que termine la tarea en curso...' 'warn'
        return
    }
    Set-Estado 'Escaneando hardware...' 'info'
    $txtScanEstado.Text = 'escaneando...'
    Reset-Ficha

    Start-Tarea -Script {
        param($corePath)
        . $corePath
        Get-InventoryRecord -IncludeSoftware $true
    } -Argumentos @($script:CorePath) -OnDone {
        param($resultado, $errorTxt)
        if ($errorTxt -or -not $resultado) {
            Set-Estado 'No se pudo leer el hardware. Ejecuta como administrador.' 'error'
            $txtScanEstado.Text = 'error'
            Write-Log "Escaneo fallido: $errorTxt" 'ERROR'
            return
        }
        $script:RegistroActual = $resultado
        $txtSerial.Text = $resultado.Serial
        $txtMarca.Text = $resultado.Marca
        $txtModelo.Text = $resultado.Modelo
        Mostrar-Ficha $resultado
        $txtScanEstado.Text = ('{0} · escaneado {1}' -f $resultado.Tipo_Equipo, (Get-Date -Format 'HH:mm:ss'))
        Set-Estado ('Equipo detectado: {0} {1}' -f $resultado.Marca, $resultado.Modelo) 'ok'
        Write-Log ('Escaneo OK: {0} {1} / {2}' -f $resultado.Marca, $resultado.Modelo, $resultado.Serial)
        $txtComentario.Focus() | Out-Null
    }
}

# =============================================================================
#  GUARDAR
# =============================================================================
function Guardar-Registro {
    param([bool]$Siguiente = $false)

    $serial = $txtSerial.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($serial) -or $serial.Length -lt 3) {
        Set-Estado 'Serial obligatorio: escanea el equipo o escribelo a mano.' 'warn'
        $txtSerial.Focus() | Out-Null
        return $false
    }

    $cliente = $txtCliente.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($cliente)) {
        Set-Estado 'Escribe el nombre del cliente antes de guardar.' 'warn'
        $txtCliente.Focus() | Out-Null
        return $false
    }

    # Construir registro: base escaneada + datos del formulario
    $reg = [ordered]@{}
    foreach ($c in $script:Columnas) {
        $reg[$c] = if ($script:RegistroActual -and $null -ne $script:RegistroActual.$c) { [string]$script:RegistroActual.$c } else { '' }
    }

    $reg['Cliente'] = $cliente
    if (-not $script:Config['Cliente']) { $script:Config['Cliente'] = $cliente; Save-Config }
    $reg['Sede_Tecnico'] = $txtSede.Text.Trim()
    $reg['Orden'] = $txtOrden.Text.Trim()
    $reg['Comentarios'] = ($txtComentario.Text -replace '[\r\n]+', ' ').Trim()
    if ($txtSerial.Text.Trim()) { $reg['Serial'] = $txtSerial.Text.Trim() }
    if ($txtMarca.Text.Trim()) { $reg['Marca'] = $txtMarca.Text.Trim() }
    if ($txtModelo.Text.Trim()) { $reg['Modelo'] = $txtModelo.Text.Trim() }

    $objeto = [PSCustomObject]$reg

    try {
        $csv = Save-RecordCsv -Registro $objeto -DataDir $script:DataDir -Cliente $cliente
        Write-Log "Registro guardado en: $(Split-Path $csv -Leaf)"
    } catch {
        Set-Estado "No se pudo escribir el CSV: $($_.Exception.Message)" 'error'
        Write-Log "Error guardando CSV: $($_.Exception.Message)" 'ERROR'
        return $false
    }

    # Cola de sincronizacion
    $archivoCola = $null
    try {
        $archivoCola = Save-PendingRecord -Registro $objeto -QueueDir $script:QueueDir -Hoja (Get-HojaActual)
        Write-Log 'Registro en cola para Google Sheets.'
    } catch {
        Write-Log "No se pudo encolar: $($_.Exception.Message)" 'WARN'
    }

    $script:ContadorSesion++
    Update-Contadores

    # Sincronizar en segundo plano si esta configurado
    if ($script:Config['SincronizarAlGuardar'] -and $script:Config['WebhookUrl']) {
        Start-Sincronizacion -Silencioso $true
    }

    Set-Estado ("Guardado: {0} ({1})" -f $serial, $cliente) 'ok'

    if ($Siguiente) {
        $script:RegistroActual = $null
        $txtSerial.Text = ''
        $txtMarca.Text = ''
        $txtModelo.Text = ''
        $txtComentario.Text = ''
        Reset-Ficha
        $txtScanEstado.Text = 'sin escanear'
        Start-Escaneo
    }

    return $true
}

# =============================================================================
#  SINCRONIZACION
# =============================================================================
function Start-Sincronizacion {
    param([bool]$Silencioso = $false)

    $url = $script:Config['WebhookUrl']
    if ([string]::IsNullOrWhiteSpace($url)) {
        if (-not $Silencioso) {
            Set-Estado 'Falta configurar WebhookUrl en config.json' 'warn'
            [System.Windows.Forms.MessageBox]::Show(
                "Para enviar a Google Sheets configura la URL del Web App en:`n`n$script:ConfigPath`n`nCampo: WebhookUrl",
                'Google Sheets', 'OK', 'Information')
        }
        return
    }

    if (-not $Silencioso) { Set-Estado 'Sincronizando con Google Sheets...' 'info' }

    # El callback se ejecuta despues; se usa una variable de script para el modo silencioso
    $script:SyncSilencioso = $Silencioso

    Start-Tarea -Script {
        param($queueDir, $url, $token, $corePath)
        . $corePath
        Sync-PendingQueue -QueueDir $queueDir -WebhookUrl $url -Token $token
    } -Argumentos @($script:QueueDir, $url, $script:Config['Token'], $script:CorePath) -OnDone {
        param($resultado, $errorTxt)
        Update-Contadores
        if ($errorTxt) {
            Write-Log "Sincronizacion fallida: $errorTxt" 'WARN'
            if (-not $script:SyncSilencioso) { Set-Estado 'No se pudo sincronizar. Se reintentara despues.' 'warn' }
            return
        }
        $enviados = if ($resultado) { $resultado.Enviados } else { 0 }
        $pendientes = if ($resultado) { $resultado.Pendientes } else { 0 }
        if ($enviados -gt 0) {
            Write-Log ("Sincronizados {0} registro(s). Pendientes: {1}" -f $enviados, $pendientes)
            Set-Estado ("Sincronizados {0} registro(s) con Google Sheets" -f $enviados) 'ok'
        } elseif ($pendientes -gt 0) {
            Set-Estado ("{0} registro(s) esperando conexion" -f $pendientes) 'warn'
        } elseif (-not $script:SyncSilencioso) {
            Set-Estado 'Todo sincronizado' 'ok'
        }
    }
}

# =============================================================================
#  ACCIONES DE SISTEMA
# =============================================================================
function Show-Mensaje {
    param([string]$Texto, [string]$Titulo = 'Wiggles Mini', [string]$Icono = 'Information')
    [System.Windows.Forms.MessageBox]::Show($Texto, $Titulo, 'OK', $Icono) | Out-Null
}

function Show-Texto {
    param([string]$Titulo, [string]$Contenido)
    $w = New-Object System.Windows.Window
    $w.Title = $Titulo
    $w.Width = 780; $w.Height = 560
    $w.WindowStartupLocation = 'CenterOwner'
    $w.Owner = $Window
    $w.Background = '#0B0F14'
    $w.Foreground = '#E6EDF3'
    $dock = New-Object System.Windows.Controls.DockPanel
    $dock.Margin = '10'

    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Text = $Contenido
    $tb.IsReadOnly = $true
    $tb.TextWrapping = 'NoWrap'
    $tb.AcceptsReturn = $true
    $tb.VerticalScrollBarVisibility = 'Auto'
    $tb.HorizontalScrollBarVisibility = 'Auto'
    $tb.Background = '#0F1520'
    $tb.Foreground = '#B9F6CA'
    $tb.BorderBrush = '#1E2A38'
    $tb.FontFamily = 'Consolas'
    $tb.FontSize = 12
    $tb.Padding = '10'
    [System.Windows.Controls.DockPanel]::SetDock($tb, 'Top')

    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = 'Cerrar'
    $btn.Width = 110
    $btn.Height = 32
    $btn.Margin = '0,10,0,0'
    $btn.Style = $Window.FindResource('BotonVerde')
    $btn.Add_Click({ $w.Close() })
    [System.Windows.Controls.DockPanel]::SetDock($btn, 'Bottom')

    $dock.Children.Add($tb) | Out-Null
    $dock.Children.Add($btn) | Out-Null
    $w.Content = $dock
    $w.ShowDialog() | Out-Null
}

function Ejecutar-Oculto {
    param([string]$Archivo, [string]$Argumentos, [string]$Descripcion)
    Write-Log "Ejecutando: $Descripcion"
    try {
        Start-Process -FilePath $Archivo -ArgumentList $Argumentos -NoNewWindow -Wait
        Write-Log "Finalizado: $Descripcion"
    } catch {
        Write-Log "Fallo $Descripcion : $($_.Exception.Message)" 'ERROR'
    }
}

function Start-Activacion {
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Se ejecutara la activacion de Windows y Office (MAS, modo silencioso).`n`nRequiere conexion a internet.`n`n¿Continuar?",
        'Activar Windows y Office', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }

    Set-Estado 'Activando Windows y Office...' 'info'
    Start-Tarea -Script {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $script = Invoke-RestMethod -Uri 'https://get.activated.win' -TimeoutSec 60
            if (-not $script) { return 'No se pudo descargar el activador (sin internet).' }
            & ([ScriptBlock]::Create($script)) /HWID /Ohook /S
            Start-Sleep -Seconds 2
            return 'OK'
        } catch {
            return "Error: $($_.Exception.Message)"
        }
    } -OnDone {
        param($resultado, $errorTxt)
        if ($resultado -eq 'OK') {
            Set-Estado 'Activacion ejecutada. Verifica el estado en Diagnostico.' 'ok'
            Write-Log 'Activacion MAS ejecutada.'
        } else {
            Set-Estado 'No se pudo activar. Revisa la conexion a internet.' 'error'
            Write-Log "Activacion fallida: $resultado $errorTxt" 'ERROR'
        }
    }
}

function Start-CambioEdicion {
    $os = Get-CimInstance Win32_OperatingSystem
    if ($os.Caption -notmatch 'Home|Core|Single') {
        Show-Mensaje "Este equipo ya es $($os.Caption).`nNo requiere cambio de edicion." 'Cambiar edicion'
        return
    }
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Se cambiara la edicion de:`n`n$($os.Caption)`n`nA: Windows Pro`n`nEl equipo se reiniciara y la activacion quedara pendiente hasta activar.`n`n¿Continuar?",
        'Cambiar edicion', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }

    Write-Log 'Iniciando cambio de edicion a Pro...'
    Set-Estado 'Cambiando edicion de Windows...' 'info'
    Start-Tarea -Script {
        try {
            $clave = 'VK7JG-NPHTM-C97JM-9MPGT-3V66T'
            $p = Start-Process -FilePath 'changepk.exe' -ArgumentList "/productkey $clave" -PassThru -Wait
            return "Codigo: $($p.ExitCode)"
        } catch { return "Error: $($_.Exception.Message)" }
    } -OnDone {
        param($resultado, $errorTxt)
        Set-Estado "Cambio de edicion: $resultado" 'ok'
        Write-Log "Cambio de edicion: $resultado"
        Show-Mensaje "Proceso finalizado ($resultado).`n`nReinicia el equipo para completar el cambio a Windows Pro." 'Cambiar edicion'
    }
}

function Start-Bloatware {
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Se eliminaran aplicaciones preinstaladas de Windows (juegos, redes sociales, apps de terceros de la Store).`n`nNo se tocan Calculadora, Fotos, Tienda ni Notas.`n`n¿Continuar?",
        'Eliminar bloatware', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }

    Set-Estado 'Eliminando bloatware...' 'info'
    Start-Tarea -Script {
        $patrones = @(
            '*CandyCrush*', '*BubbleWitch*', '*Netflix*', '*Spotify*', '*TikTok*', '*Facebook*',
            '*Instagram*', '*Disney*', '*Twitter*', '*LinkedIn*', '*Pandora*', '*Duolingo*',
            '*HiddenCity*', '*Roblox*', '*MarchofEmpires*', '*XboxApp*', '*XboxGameOverlay*',
            '*XboxGamingOverlay*', '*XboxIdentityProvider*', '*XboxSpeechToTextOverlay*',
            '*Microsoft3DViewer*', '*WindowsMaps*', '*WindowsFeedbackHub*', '*GetHelp*',
            '*GetStarted*', '*MicrosoftSolitaire*', '*BingNews*', '*BingWeather*', '*BingSports*',
            '*BingFinance*', '*ZuneMusic*', '*ZuneVideo*', '*SkypeApp*', '*OfficeHub*',
            '*OneNote*', '*People*', '*YourPhone*', '*MixedReality*', '*3DPrint*'
        )
        $eliminados = 0
        foreach ($p in $patrones) {
            Get-AppxPackage -Name $p -AllUsers -ErrorAction SilentlyContinue | ForEach-Object {
                try { Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction Stop; $eliminados++ } catch {}
            }
            Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like $p } |
                ForEach-Object {
                    try { Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction Stop | Out-Null } catch {}
                }
        }
        return $eliminados
    } -OnDone {
        param($resultado, $errorTxt)
        Set-Estado ("Bloatware eliminado ({0} paquetes)" -f $resultado) 'ok'
        Write-Log ("Bloatware: {0} paquetes eliminados" -f $resultado)
    }
}

function Start-Rendimiento {
    Set-Estado 'Configurando plan de rendimiento...' 'info'
    try {
        powercfg /setactive SCHEME_MIN 2>$null
        $activo = (powercfg /getactivescheme) -join ' '
        if ($activo -notmatch 'e9a42b02|8c5e7fda') {
            powercfg -duplicate-scheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>$null | Out-Null
        }
        powercfg /change standby-timeout-ac 0
        powercfg /change monitor-timeout-ac 20
        Set-Estado 'Plan de maximo rendimiento aplicado' 'ok'
        Write-Log 'Plan de rendimiento configurado.'
    } catch {
        Set-Estado 'No se pudo cambiar el plan de energia' 'warn'
    }
}

function Start-RepararRed {
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Se reiniciaran los componentes de red (Winsock, IP, DNS).`n`nEl equipo debe reiniciarse al terminar.`n`n¿Continuar?",
        'Reparar red', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }
    Set-Estado 'Reparando red...' 'info'
    Start-Tarea -Script {
        $salida = @()
        $salida += (netsh winsock reset 2>&1 | Out-String)
        $salida += (netsh int ip reset 2>&1 | Out-String)
        $salida += (ipconfig /flushdns 2>&1 | Out-String)
        $salida += (ipconfig /release 2>&1 | Out-String)
        $salida += (ipconfig /renew 2>&1 | Out-String)
        return ($salida -join "`n")
    } -OnDone {
        param($resultado, $errorTxt)
        Set-Estado 'Red reiniciada. Reinicia el equipo.' 'ok'
        Write-Log 'Reparacion de red ejecutada.'
    }
}

function Start-PuntoRestauracion {
    Set-Estado 'Creando punto de restauracion...' 'info'
    Start-Tarea -Script {
        try {
            Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
            Checkpoint-Computer -Description 'Wiggles Mini' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
            return 'OK'
        } catch { return "Error: $($_.Exception.Message)" }
    } -OnDone {
        param($resultado, $errorTxt)
        if ($resultado -eq 'OK') {
            Set-Estado 'Punto de restauracion creado' 'ok'
            Write-Log 'Punto de restauracion creado.'
        } else {
            Set-Estado 'No se pudo crear el punto (revisa Proteccion del sistema)' 'warn'
            Write-Log "Punto de restauracion: $resultado" 'WARN'
        }
    }
}

function Start-AdminLocal {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $usuario = [Microsoft.VisualBasic.Interaction]::InputBox('Nombre del nuevo usuario administrador:', 'Crear administrador local', 'Soporte360')
    if ([string]::IsNullOrWhiteSpace($usuario)) { return }
    $clave = [Microsoft.VisualBasic.Interaction]::InputBox('Contrasena (puede dejarla vacia):', 'Crear administrador local', '')
    try {
        if (Get-LocalUser -Name $usuario -ErrorAction SilentlyContinue) {
            Show-Mensaje "El usuario '$usuario' ya existe." 'Crear administrador'
            return
        }
        if ([string]::IsNullOrEmpty($clave)) {
            New-LocalUser -Name $usuario -NoPassword -PasswordNeverExpires -ErrorAction Stop | Out-Null
        } else {
            $segura = ConvertTo-SecureString $clave -AsPlainText -Force
            New-LocalUser -Name $usuario -Password $segura -PasswordNeverExpires -ErrorAction Stop | Out-Null
        }
        # SID bien conocido: funciona en Windows en cualquier idioma
        $grupoAdmin = Get-LocalGroup -SID 'S-1-5-32-544' -ErrorAction Stop
        Add-LocalGroupMember -Group $grupoAdmin -Member $usuario -ErrorAction Stop
        Set-Estado "Usuario '$usuario' creado como administrador" 'ok'
        Write-Log "Usuario local creado: $usuario"
    } catch {
        Set-Estado "No se pudo crear el usuario: $($_.Exception.Message)" 'error'
        Write-Log "Crear admin local: $($_.Exception.Message)" 'ERROR'
    }
}

function Show-ClaveBios {
    $clave = ''
    try { $clave = (Get-CimInstance SoftwareLicensingService).OA3xOriginalProductKey } catch {}
    if (-not $clave) {
        try { $clave = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform').BackupProductKeyDefault } catch {}
    }
    if (-not $clave) { $clave = 'No hay clave embebida en la BIOS (equipo con licencia digital o sin OEM key).' }
    Show-Texto -Titulo 'Clave de Windows' -Contenido ("Clave OEM (BIOS):`n`n  $clave`n`nY recuerda: la licencia digital queda vinculada al hardware del equipo.")
}

function Show-ResumenEquipo {
    if (-not $script:RegistroActual) {
        Show-Mensaje 'Primero escanea un equipo.' 'Resumen'
        return
    }
    $lineas = @()
    $lineas += 'FICHA DEL EQUIPO'
    $lineas += ('=' * 78)
    foreach ($c in $script:Columnas) {
        $v = $script:RegistroActual.$c
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        if ($c -eq 'Software_Lista') { $v = 'ver detalle abajo' }
        $lineas += ('{0,-24} {1}' -f ($c + ':'), $v)
    }
    if ($script:RegistroActual.Software_Lista) {
        $lineas += ''
        $lineas += 'SOFTWARE INSTALADO'
        $lineas += ('=' * 78)
        $lineas += ($script:RegistroActual.Software_Lista -replace '; ', "`n")
    }
    Show-Texto -Titulo 'Resumen del equipo' -Contenido ($lineas -join "`r`n")
}

function Start-ReporteBateria {
    Set-Estado 'Generando reporte de bateria...' 'info'
    $destino = Join-Path $script:DataDir ('Bateria_{0}_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmm'))
    try {
        powercfg /batteryreport /OUTPUT "$destino" | Out-Null
        if (Test-Path $destino) {
            Set-Estado 'Reporte de bateria generado' 'ok'
            Write-Log "Reporte de bateria: $destino"
            Start-Process $destino
        } else {
            Set-Estado 'No se pudo generar el reporte (equipo sin bateria)' 'warn'
        }
    } catch {
        Set-Estado 'No se pudo generar el reporte de bateria' 'error'
    }
}

# =============================================================================
#  OFFICE
# =============================================================================
function Get-EstadoOffice {
    $setup = Join-Path $script:OfficeDir 'Setup.exe'
    $config = Join-Path $script:OfficeDir 'Configuration.xml'

    if ((Test-Path $setup) -and (Test-Path $config)) {
        return [PSCustomObject]@{ Listo = $true; Texto = 'Listo para instalar Office 2024 desde internet (~2 GB).' }
    }
    return [PSCustomObject]@{ Listo = $false; Texto = 'Faltan Setup.exe o Configuration.xml en External\Office.' }
}

function Start-InstalarOffice {
    $setup = Join-Path $script:OfficeDir 'Setup.exe'
    $config = Join-Path $script:OfficeDir 'Configuration.xml'

    if (-not ((Test-Path $setup) -and (Test-Path $config))) {
        Set-Estado 'Faltan Setup.exe o Configuration.xml en External\Office' 'error'
        return
    }

    $r = [System.Windows.Forms.MessageBox]::Show(
        "Office 2024 se descargara de internet (~2 GB) y puede tardar bastante segun la conexion.`n`n¿Continuar?",
        'Instalar Office', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }

    Set-Estado 'Instalando Office 2024...' 'info'
    Start-Tarea -Script {
        param($officeDir)

        $setup = Join-Path $officeDir 'Setup.exe'
        $config = Join-Path $officeDir 'Configuration.xml'
        if (-not (Test-Path $setup)) { return 'Falta Setup.exe del Office Deployment Tool en External\Office' }
        if (-not (Test-Path $config)) { return 'Falta Configuration.xml en External\Office' }

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # ODT descarga el paquete y lo instala en la misma pasada
        $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$config`"" -Wait -PassThru -WindowStyle Hidden
        return [string]$p.ExitCode
    } -Argumentos @($script:OfficeDir) -OnDone {
        param($resultado, $errorTxt)
        if ($errorTxt) {
            Set-Estado 'Fallo la instalacion de Office' 'error'
            Write-Log "Office: $errorTxt" 'ERROR'
            return
        }
        if ($resultado -eq '0') {
            Set-Estado 'Office instalado correctamente' 'ok'
            Write-Log 'Office 2024 instalado.'
        } elseif ($resultado -match '^Falta') {
            Set-Estado $resultado 'error'
            Write-Log $resultado 'ERROR'
        } else {
            Set-Estado "Office finalizo con codigo $resultado" 'warn'
            Write-Log "Office exit code: $resultado" 'WARN'
        }
        Actualizar-EstadoOffice
    }
}

function Actualizar-EstadoOffice {
    $e = Get-EstadoOffice
    $txtOfficeEstado.Text = $e.Texto
    $txtOfficeEstado.Foreground = if ($e.Listo) { '#B9F6CA' } else { '#FFB300' }
}

function Start-QuitarOffice {
    $r = [System.Windows.Forms.MessageBox]::Show(
        "Se desinstalara cualquier version de Office instalada (Click-to-Run).`n`n¿Continuar?",
        'Desinstalar Office', 'YesNo', 'Warning')
    if ($r -ne 'Yes') { return }
    Set-Estado 'Desinstalando Office...' 'info'
    Start-Tarea -Script {
        $setup = Join-Path $env:ProgramFiles 'Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe'
        if (-not (Test-Path $setup)) {
            $setup = Join-Path ${env:ProgramFiles(x86)} 'Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe'
        }
        if (-not (Test-Path $setup)) { return 'No se encontro Office Click-to-Run instalado.' }
        $p = Start-Process -FilePath $setup -ArgumentList 'scenario=install scenario=uninstall' -Wait -PassThru
        return "Desinstalacion finalizada (codigo $($p.ExitCode))."
    } -OnDone {
        param($resultado, $errorTxt)
        Set-Estado 'Proceso de desinstalacion terminado' 'ok'
        Write-Log "Quitar Office: $resultado"
    }
}

# =============================================================================
#  SOFTWARE
# =============================================================================
function Start-AppsBasicas {
    $lista = @($script:Config['AppsBasicas'])
    if ($lista.Count -eq 0) { Show-Mensaje 'No hay aplicaciones configuradas en config.json' 'Aplicaciones'; return }

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Show-Mensaje 'Winget no esta disponible en este equipo. Usa "Reparar winget" primero.' 'Aplicaciones' 'Warning'
        return
    }

    Set-Estado 'Instalando aplicaciones basicas...' 'info'
    Write-Log ("Instalando: " + ($lista -join ', '))
    Start-Tarea -Script {
        param($ids)
        $ok = 0; $fail = 0
        foreach ($id in $ids) {
            try {
                $p = Start-Process -FilePath 'winget' -ArgumentList @(
                    'install', '--id', $id, '-e', '--silent',
                    '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity'
                ) -Wait -PassThru -WindowStyle Hidden
                if ($p.ExitCode -eq 0) { $ok++ } else { $fail++ }
            } catch { $fail++ }
        }
        return [PSCustomObject]@{ Ok = $ok; Fallidos = $fail }
    } -Argumentos @(, $lista) -OnDone {
        param($resultado, $errorTxt)
        if ($resultado) {
            Set-Estado ("Aplicaciones instaladas: {0} · fallidas: {1}" -f $resultado.Ok, $resultado.Fallidos) 'ok'
            Write-Log ("Apps basicas: {0} OK / {1} fallidas" -f $resultado.Ok, $resultado.Fallidos)
        } else {
            Set-Estado 'No se pudieron instalar las aplicaciones' 'error'
        }
    }
}

function Start-Runtimes {
    Set-Estado 'Instalando runtimes...' 'info'
    Start-Tarea -Script {
        $ids = @(
            'Microsoft.VCRedist.2015+.x64',
            'Microsoft.VCRedist.2015+.x86',
            'Microsoft.DotNet.DesktopRuntime.8'
        )
        foreach ($id in $ids) {
            try {
                Start-Process -FilePath 'winget' -ArgumentList @(
                    'install', '--id', $id, '-e', '--silent',
                    '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity'
                ) -Wait -WindowStyle Hidden
            } catch {}
        }
        return 'OK'
    } -OnDone {
        param($resultado, $errorTxt)
        Set-Estado 'Runtimes procesados' 'ok'
    }
}

function Start-RepararWinget {
    Set-Estado 'Reparando winget...' 'info'
    Start-Tarea -Script {
        try {
            $destino = Join-Path $env:TEMP 'winget.msixbundle'
            Invoke-WebRequest -Uri 'https://aka.ms/getwinget' -OutFile $destino -UseBasicParsing -TimeoutSec 180
            Add-AppxPackage -Path $destino -ErrorAction Stop
            $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
            Start-Process -FilePath 'winget' -ArgumentList 'source update --accept-source-agreements --disable-interactivity' -Wait -WindowStyle Hidden
            return 'OK'
        } catch { return "Error: $($_.Exception.Message)" }
    } -OnDone {
        param($resultado, $errorTxt)
        if ($resultado -eq 'OK') {
            Set-Estado 'Winget reparado' 'ok'
            Write-Log 'Winget reparado.'
        } else {
            Set-Estado 'No se pudo reparar winget' 'error'
            Write-Log "Winget: $resultado" 'ERROR'
        }
    }
}

# =============================================================================
#  HERRAMIENTAS PORTABLES
# =============================================================================
$script:Tools = @(
    @{ Nombre = '3DP Net';          Archivo = '3DP_Net*.exe';           Desc = 'Drivers de red offline (equipos sin internet)' }
    @{ Nombre = '3DP Chip';         Archivo = '3DP_Chip*.exe';          Desc = 'Detecta y descarga drivers del equipo' }
    @{ Nombre = 'CrystalDiskInfo';  Archivo = 'CrystalDiskInfo*.exe';   Desc = 'Salud SMART del disco' }
    @{ Nombre = 'CrystalDiskMark';  Archivo = 'CrystalDiskMark*.exe';   Desc = 'Velocidad del disco' }
    @{ Nombre = 'HWiNFO';           Archivo = 'HWiNFO*.exe';            Desc = 'Sensores y temperatura' }
    @{ Nombre = 'UninstallTool';    Archivo = 'UninstallTool*.exe';     Desc = 'Desinstalacion limpia' }
)

function Construir-PanelTools {
    $pnlTools.Children.Clear()
    $encontradas = 0

    foreach ($t in $script:Tools) {
        $ruta = $null
        if (Test-Path $script:ToolsDir) {
            try {
                $ruta = Get-ChildItem -Path $script:ToolsDir -Filter $t.Archivo -Recurse -File -ErrorAction Ignore |
                    Sort-Object Length -Descending | Select-Object -First 1
            } catch {}
        }

        $card = New-Object System.Windows.Controls.Border
        $card.Style = $Window.FindResource('Tarjeta')
        $card.Width = 300
        $card.Height = 118

        $sp = New-Object System.Windows.Controls.StackPanel
        $titulo = New-Object System.Windows.Controls.TextBlock
        $titulo.Text = $t.Nombre
        $titulo.FontWeight = 'SemiBold'
        $titulo.FontSize = 13
        $titulo.Foreground = '#E6EDF3'
        $sp.Children.Add($titulo) | Out-Null

        $desc = New-Object System.Windows.Controls.TextBlock
        $desc.Text = $t.Desc
        $desc.FontSize = 10.5
        $desc.Foreground = '#8494A6'
        $desc.TextWrapping = 'Wrap'
        $desc.Margin = '0,2,0,8'
        $sp.Children.Add($desc) | Out-Null

        $btn = New-Object System.Windows.Controls.Button
        if ($ruta) {
            $btn.Content = 'ABRIR'
            $btn.Style = $Window.FindResource('BotonVerde')
            $btn.Height = 30
            $btn.Tag = $ruta.FullName
            $btn.Add_Click({
                param($s, $e)
                $exe = $s.Tag
                Write-Log "Abriendo herramienta: $exe"
                try { Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe -Parent) } catch {
                    Show-Mensaje "No se pudo abrir: $exe" 'Herramienta' 'Warning'
                }
            })
            $encontradas++
        } else {
            $btn.Content = 'NO INSTALADA'
            $btn.IsEnabled = $false
            $btn.Height = 30
            $btn.Tag = $t.Archivo
            $btn.ToolTip = "Coloca el archivo en External\Tools\$($t.Archivo)"
        }
        $sp.Children.Add($btn) | Out-Null

        $card.Child = $sp
        $pnlTools.Children.Add($card) | Out-Null
    }

    $txtToolsInfo.Text = if ($encontradas -eq 0) {
        "No hay herramientas en External\Tools. Copia ahi las utilidades portables (CrystalDiskInfo, HWiNFO, 3DP Net, UninstallTool)."
    } else {
        "$encontradas de $($script:Tools.Count) herramientas disponibles en el pendrive."
    }
}

# =============================================================================
#  EXPORTAR
# =============================================================================
function Start-ExportarExcel {
    Set-Estado 'Generando Excel...' 'info'
    $cliente = $txtCliente.Text.Trim()
    $csv = Get-CsvPath -DataDir $script:DataDir -Cliente $cliente
    if (-not (Test-Path $csv)) {
        $csv = Get-ChildItem -Path $script:DataDir -Filter 'Inventario_*.csv' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $csv -or -not (Test-Path $csv)) {
        Set-Estado 'Todavia no hay registros para exportar' 'warn'
        return
    }

    $xlsx = [System.IO.Path]::ChangeExtension($csv, '.xlsx')
    try {
        Export-CsvToXlsx -CsvPath $csv -XlsxPath $xlsx | Out-Null
        Set-Estado "Excel generado: $(Split-Path $xlsx -Leaf)" 'ok'
        Write-Log "Excel exportado: $xlsx"
        Start-Process $xlsx
    } catch {
        Set-Estado "No se pudo generar el Excel: $($_.Exception.Message)" 'error'
        Write-Log "Exportar Excel: $($_.Exception.Message)" 'ERROR'
    }
}

# =============================================================================
#  EVENTOS
# =============================================================================
$btnMin.Add_Click({ $Window.WindowState = 'Minimized' })
$btnCerrar.Add_Click({ $Window.Close() })

$btnEscanear.Add_Click({ Start-Escaneo })
$btnGuardar.Add_Click({ Guardar-Registro -Siguiente $false | Out-Null })
$btnGuardarSiguiente.Add_Click({ Guardar-Registro -Siguiente $true | Out-Null })

$btnExcel.Add_Click({ Start-ExportarExcel })
$btnExportarTodo.Add_Click({ Start-ExportarExcel })
$btnSincronizar.Add_Click({ Start-Sincronizacion })
$btnSyncTodo.Add_Click({ Start-Sincronizacion })

$btnCarpeta.Add_Click({
    try { Start-Process explorer.exe $script:DataDir } catch {}
})

$btnActivar.Add_Click({ Start-Activacion })
$btnEdicion.Add_Click({ Start-CambioEdicion })
$btnBloatware.Add_Click({ Start-Bloatware })
$btnRendimiento.Add_Click({ Start-Rendimiento })
$btnRed.Add_Click({ Start-RepararRed })
$btnPuntoRestauracion.Add_Click({ Start-PuntoRestauracion })
$btnAdminLocal.Add_Click({ Start-AdminLocal })
$btnClaveBios.Add_Click({ Show-ClaveBios })
$btnReporteBateria.Add_Click({ Start-ReporteBateria })
$btnResumenEquipo.Add_Click({ Show-ResumenEquipo })

$btnOffice.Add_Click({ Start-InstalarOffice })
$btnQuitarOffice.Add_Click({ Start-QuitarOffice })
$btnApps.Add_Click({ Start-AppsBasicas })
$btnRuntimes.Add_Click({ Start-Runtimes })
$btnWingetReparar.Add_Click({ Start-RepararWinget })

# Atajos de teclado
$Window.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'F5') { Start-Escaneo; $e.Handled = $true }
    if ($e.Key -eq 'F8') { Guardar-Registro -Siguiente $true | Out-Null; $e.Handled = $true }
})

$txtSerial.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'Return') { $txtComentario.Focus() | Out-Null }
})

# =============================================================================
#  ARRANQUE
# =============================================================================
$Window.Add_Loaded({
    Write-Log ("Wiggles Mini {0} iniciado en {1}" -f $script:AppVersion, $script:Root)
    Write-Log ("Administrador: {0} · Equipo: {1}" -f (Test-EsAdmin), $env:COMPUTERNAME)

    if ($script:Config['Cliente']) { $txtCliente.Text = $script:Config['Cliente'] }
    if ($script:Config['Tecnico']) { $txtSede.Text = $script:Config['Tecnico'] }
    $txtOrden.Text = ''
    $txtAppsLista.Text = (@($script:Config['AppsBasicas']) -join ' · ')

    Actualizar-EstadoOffice
    Construir-PanelTools
    Update-Contadores

    # Estado de licencias en segundo plano
    Start-Tarea -Script {
        param($corePath)
        . $corePath
        $lic = Get-EstadoLicencia
        return $lic
    } -Argumentos @($script:CorePath) -OnDone {
        param($lic, $errorTxt)
        if ($lic) {
            $txtLicencia.Text = "Windows: {0}   |   Office: {1}" -f $lic.Windows, $lic.Office
            $txtLicencia.Foreground = if ($lic.Windows -eq 'Activado') { '#B9F6CA' } else { '#FFB300' }
        }
    }

    Set-Estado 'Listo. Escanea el primer equipo.' 'ok'
})

$Window.Add_Closing({
    Save-Config
    Write-Log 'Sesion finalizada.'
})

# Limpieza de tareas al cerrar
$Window.Add_Closed({
    foreach ($t in @($script:Tareas)) {
        try { $t.PS.Stop() } catch {}
        try { $t.PS.Dispose() } catch {}
        try { $t.RS.Dispose() } catch {}
    }
})

# Arrancar escaneo automatico del primer equipo (espera a que terminen las tareas iniciales)
$Window.Add_ContentRendered({
    $script:TimerEscaneoInicial = New-Object System.Windows.Threading.DispatcherTimer
    $script:TimerEscaneoInicial.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:TimerEscaneoInicial.Add_Tick({
        if ($script:TareasActivas -gt 0) { return }
        $script:TimerEscaneoInicial.Stop()
        Start-Escaneo
    })
    $script:TimerEscaneoInicial.Start()
})

[void]$Window.ShowDialog()
