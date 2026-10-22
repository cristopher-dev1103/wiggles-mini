# Wiggles Mini

Inventario de equipos para visitas a clientes. Portable, sin instalación y sin dependencias:
un `.bat`, dos `.ps1`, catálogos JSON y una carpeta `External`. Guarda siempre en local (CSV)
y sincroniza con Google Sheets cuando hay internet. Incluye módulos portados de WinUtil
(aplicaciones, ajustes, reparaciones y apps preinstaladas) con interfaz propia en español.

![flujo](https://img.shields.io/badge/PowerShell-5.1%2B-blue) ![windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6)

---

## 1. Qué hace

| Módulo | Detalle |
|---|---|
| **Inventario** | 70 campos: equipo, CPU, RAM (módulos y slots), discos y SMART, batería (desgaste y ciclos), SO, licencias, red, TPM, Secure Boot, antivirus y software instalado. |
| **Google Sheets** | Una pestaña nueva por visita (`2026-10-22 · Cliente`) en el libro maestro. |
| **Office** | Instala Office 2024 con el ODT. Descarga el paquete de Microsoft durante la instalación. |
| **Aplicaciones** | 251 programas con búsqueda y 15 categorías; instala con winget o Microsoft Store (7-Zip, Chrome, PDF, utilidades, juegos, etc.). |
| **Ajustes** | 33 ajustes de privacidad, interfaz, sistema y mantenimiento. Guarda el valor original y permite deshacer. |
| **Config** | 8 reparaciones (SFC, Windows Update, red, NTP, temporales, DISM, winget, restauración), 7 características de Windows y 14 paneles clásicos. |
| **AppX** | 34 apps preinstaladas de Windows: quitar, reinstalar y detectar instaladas. |
| **Activación** | Windows y Office (módulo existente de la herramienta original). |
| **Edición** | Cambio de Windows Home a Pro con la clave genérica de Microsoft. |
| **Utilidades** | Punto de restauración, bloatware, plan de máximo rendimiento, reparar red, usuario administrador local, batería HTML, clave OEM del BIOS. |

Todo se ejecuta desde el pendrive. No escribe en el equipo del cliente salvo lo que
apliques a propósito (Office, aplicaciones, ajustes o reparaciones).

---

## 2. Contenido

```
2-Wiggles_Mini\
├─ INICIAR.bat            ← doble clic aquí
├─ WigglesMini.ps1        ← interfaz (WPF embebido)
├─ WigglesCore.ps1        ← escaneo, CSV, cola y Excel
├─ config.json            ← cliente, técnico, webhook y token
├─ apps.json              ← catálogo de 251 aplicaciones (winget / msstore)
├─ appx.json              ← catálogo de 34 apps preinstaladas
├─ tweaks.json            ← catálogo de 33 ajustes con deshacer
├─ irm.ps1                ← arranque remoto (irm | iex)
├─ README.md
├─ TERCEROS.md            ← atribución de WinUtil y licencia MIT
├─ Publicar-GitHub.md     ← cómo publicarlo y usar dominio propio
├─ Sheets\
│  └─ Code.gs             ← Apps Script + instalador del libro
├─ External\
│  └─ Office\             ← Setup.exe (ODT) + Configuration.xml
└─ Data\                  ← se crea sola
   ├─ Inventario_<Cliente>_<fecha>.csv
   ├─ Inventario_<Cliente>_<fecha>.xlsx
   ├─ Ajustes\            ← estado.json de ajustes aplicados
   ├─ Pendientes\         ← registros sin enviar
   └─ Logs\
```

Tamaño: **~1 MB** (incluye el ODT de Office). Office se descarga aparte durante la instalación.
Para llevarlo en un pendrive basta copiar la carpeta completa.

---

## 3. Uso en 4 pasos

1. Copia la carpeta al pendrive.
2. Doble clic en `INICIAR.bat`. Pide permisos de administrador (necesario para TPM y Secure Boot).
3. Escribe **Cliente** y **Sede/Técnico**. El serial y el hardware se llenan solos al abrir.
4. **GUARDAR Y SIGUIENTE EQUIPO** graba el CSV, encola el envío y limpia el formulario para el próximo equipo.

Al terminar la visita pulsa **Sincronizar** (o deja que se sincronice al guardar si hay internet).
Si no hubo red, los registros quedan en `Data\Pendientes` y se envían solos en la siguiente
oportunidad, sin duplicarse.

> El botón **Exportar Excel** genera un `.xlsx` real (con filtro y encabezado) desde el CSV del día.
> **Abrir carpeta** lleva a `Data\`.
>
> ¿Sin pendrive? También puedes ejecutarlo con `irm | iex` (ver sección 10).

---

## 4. Configuración (`config.json`)

```json
{
  "Cliente": "",
  "Tecnico": "",
  "WebhookUrl": "https://script.google.com/macros/s/XXXX/exec",
  "Token": "WIGGLES-CAMBIAR-ESTE-TOKEN",
  "AppsBasicas": ["7zip.7zip", "Google.Chrome", "VideoLAN.VLC"],
  "SincronizarAlGuardar": true
}
```

| Campo | Uso |
|---|---|
| `Cliente` | Se recuerda entre sesiones; se puede cambiar en la interfaz. |
| `Tecnico` | Se carga en *Sede / Técnico*. |
| `WebhookUrl` | URL `/exec` de la implementación de Apps Script. |
| `Token` | Debe ser idéntico al de `Code.gs`. Si está vacío, no se envía nada. |
| `AppsBasicas` | IDs de winget para el botón **Instalar aplicaciones básicas**. |
| `SincronizarAlGuardar` | `true` intenta enviar al guardar; `false` solo con el botón. |

Los cambios hechos en la interfaz se guardan solos al cerrar.

---

## 5. Google Sheets (el script crea todo)

`Sheets\Code.gs` incluye un instalador: prepara el libro maestro, la pestaña **INICIO** y
el **token**. No hay que crear pestañas a mano.

1. Entra a <https://script.google.com> → **Nuevo proyecto**.
2. Borra el contenido y pega `Sheets\Code.gs`.
3. Arriba elige la función **`instalar`** y pulsa **Ejecutar**. Autoriza los permisos.
   - Usa el libro indicado en `SPREADSHEET_ID` (ya viene configurado).
   - Genera el token y lo escribe en la pestaña **INICIO**.
   - El token también queda en el **registro de ejecución**.
4. **Implementar → Nueva implementación → Aplicación web**
   - *Ejecutar como:* Yo
   - *Quién tiene acceso:* Cualquier persona
5. Copia la URL que termina en `/exec` y el token en `WebhookUrl` y `Token` de `config.json`.
6. Prueba: abre la URL en el navegador. Debe responder `{"ok":true,...}`.

> Si prefieres usar un libro existente, pon su ID en `SPREADSHEET_ID` antes de ejecutar
> `instalar()`; en ese caso no crea uno nuevo.

Funciones disponibles en el editor:

| Función | Para qué |
|---|---|
| `instalar` | Crea libro + pestaña INICIO + token |
| `mostrarToken` | Imprime el token actual en el registro |
| `rotarToken` | Genera un token nuevo y actualiza INICIO |
| `estado` | Muestra libro, webhook y pestañas |

Cada visita crea su pestaña con encabezado verde, fila fija, filtro y formato.
Si un registro se reenvía, el `id` evita duplicarlo.

> Cada vez que edites `Code.gs` usa **Implementar → Gestionar implementaciones → Editar**
> para conservar la misma URL. Si creas una implementación nueva, cambia la URL en `config.json`.

---

## 6. Office 2024

- `External\Office\Setup.exe` es el **Office Deployment Tool** y `Configuration.xml` define
  Office Profesional Plus 2024 en español, sin interfaz.
- **Instalar Office 2024** descarga el paquete (~2 GB) y lo instala en la misma pasada.
  Requiere conexión a internet estable en el equipo del cliente.
- El pendrive no guarda el paquete: no crece y siempre se instala la versión vigente.

Para cambiar de edición o idioma edita `Configuration.xml` (producto, `Language`, `ExcludeApp`).

La activación de Windows/Office usa el módulo de la herramienta original y requiere internet
y reinicio al finalizar.

---

## 7. Aplicaciones (winget)

La pestaña **APLICACIONES** trae un catálogo de 251 programas agrupados en 15 categorías
(la base inicial de 75 se amplió con el catálogo de WinUtil):

- Busca por nombre, descripción o id de winget.
- Filtra por categoría (Navegadores, Oficina, Utilidades, Multimedia, Juegos, etc.).
- Las marcadas con **★** son recomendadas; **SELECCIONAR RECOMENDADAS** las marca todas.
- **INSTALAR SELECCIONADAS** instala en silencio, una por una, y muestra cuáles fallaron.
- Soporta ids de winget y de Microsoft Store (`msstore:...`); los de Store se instalan con
  `--source msstore`.

> Las 75 apps originales conservan sus descripciones en español; las nuevas usan una
> descripción breve por categoría. El **Id** es el dato que realmente usa winget.

No se descargan instaladores de portales como FilePuma: winget usa el manifiesto
oficial del fabricante y valida el hash. Si un programa no está en winget ni en Microsoft
Store, no aparece en el catálogo.

El catálogo está en `apps.json`. Para agregar uno, usa el **Id** de winget:

```json
{ "Id": "MiPrograma.Id", "Nombre": "Mi Programa", "Categoria": "Utilidades", "Descripcion": "Que hace", "Recomendada": false }
```

Busca el Id con:

```powershell
winget search "nombre del programa"
```

Requiere conexión a internet. Si winget falla, usa **Reparar winget / App Installer**.
El botón **APLICACIONES BASICAS** de la pestaña SOFTWARE sigue usando la lista corta de
`config.json` (`AppsBasicas`).

---

## 8. Ajustes, Config y AppX

Estos tres módulos se portaron desde WinUtil (datos y definiciones) y se reescribieron
con una interfaz y un motor propios en español.

### AJUSTES

- 33 tarjetas agrupadas en **Privacidad**, **Interfaz**, **Sistema** y **Mantenimiento**.
- Cada tarjeta indica **SEGURO** o **AVANZADO**. Los avanzados pueden afectar funciones
  del sistema; aplícalos solo si sabes lo que hacen.
- **APLICAR** guarda el valor actual (registro, servicios y tipo de inicio) en
  `Data\Ajustes\estado.json`; **DESHACER** lo restaura.
- Los ajustes con script (por ejemplo, reiniciar Explorer o cambiar la hibernación) también
  ejecutan su script de deshacer.
- **Widgets** se puede aplicar pero no deshacer desde la herramienta; se marca como
  `APLICADO (SIN DESHACER)`.
- **Limpieza de componentes** y **Archivos temporales** son acciones puntuales: se pueden
  ejecutar varias veces y no guardan estado.
- Antes de aplicar varios ajustes en un equipo de cliente, crea un **punto de restauración**
  desde SISTEMA. El estado de ajustes no sustituye un respaldo del sistema.

### CONFIG

**Reparaciones**

| Botón | Qué hace |
|---|---|
| SFC + DISM | `sfc /scannow` y `DISM /RestoreHealth`. Puede tardar 20-40 minutos. |
| Resetear Windows Update | Detiene servicios, renombra `SoftwareDistribution` y `Catroot2`, reregistra DLLs y reinicia servicios. No borra políticas de grupo. |
| Resetear red | `winsock reset`, `int ip reset`, `winhttp reset proxy`, flush/register DNS. Requiere reinicio. |
| Hora NTP | Configura `pool.ntp.org` y resincroniza el reloj. |
| Limpiar temporales | Borra temporales del usuario y de Windows. |
| Limpiar componentes (DISM) | `DISM /StartComponentCleanup` (sin `ResetBase`). |
| Reparar winget | Reinstala/registra App Installer y winget. |
| Punto de restauración | Crea un punto de restauración del sistema. |

**Características de Windows**: .NET Framework 2/3/4, Hyper-V, WSL, Sandbox, cliente NFS,
reproductor/DirectPlay y servidor OpenSSH. Se instalan con
`Enable-WindowsOptionalFeature`. Hyper-V y Sandbox no están disponibles en Windows Home.
Puede requerir reinicio y descarga desde Windows Update.

**Paneles clásicos**: 14 accesos directos (Administración de equipos, Panel de control,
mouse, red, energía, impresoras, programas, región, seguridad, sonido, sistema, fecha/hora,
firewall y restauración).

### APPX

- 34 apps preinstaladas, agrupadas por categoría.
- **PREDETERMINADOS** marca 13 apps de bloatware común (Bing, Copilot, Noticias, Clima,
  Xbox, Dev Home, Power Automate, Solitaire, etc.). Calculadora, Fotos, Notepad, Paint,
  Cámara y Recortes no se marcan por defecto.
- **QUITAR SELECCIONADAS** elimina el paquete del usuario actual y, si existe, el paquete
  provisionado para nuevos usuarios.
- **INSTALAR SELECCIONADAS** intenta registrar el manifiesto local; si no está disponible,
  usa `winget --source msstore` con el `StoreId`.
- **DETECTAR INSTALADAS** marca con `●` las apps presentes. Al abrir la pestaña se ejecuta
  una detección automática.
- Quitar una app puede afectar dependencias (por ejemplo, Xbox o Store). Si algo falla,
  vuelve a instalar desde Microsoft Store.

> Los catálogos `apps.json`, `appx.json` y `tweaks.json` derivan de Chris Titus WinUtil
> (MIT). Consulta `TERCEROS.md` para la atribución y la licencia.

---

## 9. Problemas frecuentes

| Síntoma | Causa / solución |
|---|---|
| `... no se puede cargar porque la ejecución de scripts está deshabilitada` | Usa `INICIAR.bat` (aplica `-ExecutionPolicy Bypass`). |
| TPM y Secure Boot dicen `Requiere admin` | Abre con `INICIAR.bat` y acepta el UAC. |
| No sincroniza | Revisa `WebhookUrl`, `Token` y que la implementación sea *Cualquier persona*. Los registros quedan en `Data\Pendientes`. |
| Sheet responde `Token invalido` | Ejecuta `mostrarToken()` en Apps Script y copia ese valor en `config.json`. |
| No aparece la pestaña | Abre la pestaña **INICIO** del libro que creó `instalar()`; ahí está la URL real. |
| Office no instala | Verifica la conexión a internet y que `Setup.exe` y `Configuration.xml` estén en `External\Office`. |
| Un ajuste quedó a medias | Revisa `Data\Logs\Mini_<fecha>.log`. Pulsa **DESHACER** en la tarjeta; si no aparece, borra su entrada de `Data\Ajustes\estado.json` solo después de restaurar el valor a mano. |
| Una app AppX no se reinstala | Necesita internet y Microsoft Store. Si no tiene `StoreId`, solo se puede registrar el manifiesto local; revisa el log. |
| Windows Update sigue fallando | El reset moderado no borra políticas de grupo. Revisa `gpedit`/registro o usa el punto de restauración. |
| Las apps nuevas tienen descripción genérica | El catálogo se amplió desde WinUtil; el **Id** y el **Nombre** son los datos importantes. Puedes editar `Descripcion` en `apps.json`. |
| Antivirus bloquea el `.bat` | Es un lanzador de PowerShell sin firma; agrega la carpeta a exclusiones. |
| Acentos raros en los `.ps1` | Guárdalos siempre como **UTF-8 con BOM** si los editas. |

Los detalles de cada sesión quedan en `Data\Logs\Mini_<fecha>.log`.

---

## 10. Modo remoto (sin pendrive)

Pensado para cuando se olvida el pendrive. Jeric solo pega:

```powershell
irm "https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1" | iex
```

Enlaces cortos (más fáciles de recordar):

```powershell
irm "https://is.gd/wiggles" | iex
```

Respaldo: `https://tinyurl.com/wiggles-irm`. Ambos redirigen al `irm.ps1` del
repositorio; si algún día cambias de repositorio, crea un alias nuevo.

`irm.ps1` descarga la herramienta a `%TEMP%\WigglesMini`, pide administrador y la abre.
Siempre trae la última versión publicada.

> En modo remoto los CSV y la cola quedan en `%TEMP%\WigglesMini\Data`. Si el equipo se
> apaga o se limpia el temporal, se pierden; sincroniza antes de cerrar.

Requisitos y pasos completos (incluido dominio propio tipo `tudominio.com/irm`):
**`Publicar-GitHub.md`**.

> El modo remoto no guarda el paquete de Office; la instalación de Office sigue siendo
> por internet, igual que con el pendrive.

---

## 11. Notas

- Windows 10/11, PowerShell 5.1 (incluido). No requiere módulos ni instalar nada.
- No se envía información a ningún sitio salvo el webhook que configures.
- La activación reutiliza el módulo de la herramienta original; úsala solo en equipos
  con licencia válida.
