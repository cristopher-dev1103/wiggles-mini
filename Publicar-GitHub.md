# Publicar Wiggles Mini en GitHub (`irm | iex`)

Con esto Jeric puede trabajar sin el pendrive:

```powershell
irm "https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1" | iex
```

El script descarga la herramienta a `%TEMP%\WigglesMini`, pide administrador y la abre.

---

## 1. Crear el repositorio

1. En GitHub: **New repository**.
2. Nombre: `wiggles-mini`.
3. Visibilidad: **Public** (si es privado, `irm` no puede leerlo sin un proxy; ver punto 9).
4. Marca **Add a README** si quieres, o créalo vacío.

## 2. Subir los archivos

Sube **el contenido** de `2-Wiggles_Mini`, no la carpeta padre:

```
irm.ps1
WigglesMini.ps1
WigglesCore.ps1
config.json
apps.json
README.md
Sheets/Code.gs
External/Office/Setup.exe
External/Office/Configuration.xml
```

Opciones:

- **Web**: *Add file → Upload files*, arrastra todo y confirma en `main`.
- **Git**:

```powershell
cd "C:\Users\c.rojas\Documents\PROYECTOS_360\REPORTE\2-Wiggles_Mini"
git init
git add .
git commit -m "Wiggles Mini 1.0"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/wiggles-mini.git
git push -u origin main
```

> `Setup.exe` pesa 660 KB, así que no hay problema con el límite de GitHub.

## 3. Configurar la URL base

Edita `irm.ps1` **en el repositorio** (o antes de subirlo) y cambia:

```powershell
$BaseUrl = 'https://raw.githubusercontent.com/TU-USUARIO/wiggles-mini/main'
```

por tu usuario real. Si no lo cambias, el script se detiene con un error claro.

## 4. Completar `config.json`

Antes de publicar, deja listos el webhook y el token del libro:

```json
"WebhookUrl": "https://script.google.com/macros/s/XXXX/exec",
"Token": "el-token-que-te-dio-instalar()"
```

Así Jeric solo pega el comando y no configura nada.

> Si el repositorio es público, el token queda visible. Mitigaciones en el punto 8.

## 5. Probar

En un PowerShell normal:

```powershell
irm "https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1" | iex
```

Debe mostrar la descarga, pedir UAC y abrir la interfaz. Si cambias archivos en el
repositorio, el siguiente `irm` ya trae la versión nueva (GitHub cachea ~5 minutos).

## 6. Actualizar

- Edita el archivo en GitHub y confirma el cambio.
- No hace falta recompilar nada: `irm` descarga siempre los `.ps1` actuales.

## 7. Dominio propio (`tudominio.com/irm`) — opcional

Para que quede como el de Titus necesitas un dominio propio y algo que sirva el texto.
La forma más simple es **Cloudflare Workers** (plan gratuito):

> `winscript.cc` ya está ocupado por otro proyecto público (Winscript); usa tu
> propio dominio.

1. Agrega tu dominio a Cloudflare y apunta el DNS a Cloudflare.
2. Crea un Worker con este código:

```js
export default {
  async fetch(request) {
    const url = 'https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1';
    const res = await fetch(url, { cf: { cacheTtl: 0 } });
    return new Response(res.body, {
      headers: {
        'content-type': 'text/plain; charset=utf-8',
        'cache-control': 'no-store'
      }
    });
  }
};
```

3. En el Worker: **Settings → Triggers → Custom Domains → Add** `tudominio.com`
   (o una ruta `tudominio.com/irm*`).
4. Prueba:

```powershell
irm "https://tudominio.com/irm" | iex
```

> `irm` necesita que la respuesta sea texto plano. Un redirect 302 también sirve:
> puedes crear una regla que mande `/irm` al raw de GitHub.

### Enlace corto sin dominio propio

Si todavía no tienes dominio, usa un acortador que redirija directo (sin página
intermedia). El proyecto ya tiene dos:

```powershell
irm "https://is.gd/wiggles" | iex
```

Respaldo: `https://tinyurl.com/wiggles-irm`.

Ventaja: funciona en segundos. Desventaja: depende de un servicio externo y el
alias no se puede editar si lo creaste sin cuenta. Para algo permanente y con tu
marca, usa un dominio propio con Cloudflare.

Para crear otro alias (sin registro):

```text
https://is.gd/create.php?format=simple&url=URL_LARGA&shorturl=ALIAS
https://tinyurl.com/api-create.php?url=URL_LARGA&alias=ALIAS
```

## 8. Seguridad del token

El token solo permite **agregar filas** al libro, no leer ni borrar. Aun así:

- Si el repo es público, cualquiera que lo vea puede escribir en tu hoja.
- Rota el token cuando haga falta: ejecuta `rotarToken()` en Apps Script y actualiza
  `config.json` en el repositorio.
- Si quieres ocultarlo, usa **repositorio privado + Worker** con un secreto de GitHub:

```js
export default {
  async fetch(request, env) {
    const res = await fetch(
      'https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1',
      { headers: { authorization: `Bearer ${env.GITHUB_TOKEN}` } }
    );
    return new Response(res.body, {
      headers: { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-store' }
    });
  }
};
```

y guarda `GITHUB_TOKEN` en **Settings → Variables and Secrets** del Worker.

## 9. Alternativa sin publicar el token en el repo

Si prefieres no tocar `config.json` en GitHub, deja el webhook vacío y usa variables
de entorno antes del `irm`:

```powershell
$env:WIGGLES_URL = 'https://script.google.com/macros/s/XXXX/exec'
$env:WIGGLES_TOKEN = 'tu-token'
irm "https://raw.githubusercontent.com/cristopher-dev1103/wiggles-mini/main/irm.ps1" | iex
```

`irm.ps1` las aplica sobre el `config.json` descargado en `%TEMP%\WigglesMini`.

---

## Resumen del flujo

| Paso | Quién | Qué hace |
|---|---|---|
| 1 | Tú | Creas el libro y el token con `instalar()` en Apps Script. |
| 2 | Tú | Publicas el repo y pones la URL/token en `config.json`. |
| 3 | Jeric | `irm ".../irm.ps1" \| iex` en cada equipo. |
| 4 | Ambos | Revisan la pestaña de la visita en Google Sheets. |
