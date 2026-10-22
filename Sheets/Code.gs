/**
 * WIGGLES MINI - RECEPTOR DE INVENTARIO
 * -------------------------------------
 * Recibe cada equipo escaneado por WigglesMini.ps1 y lo agrega como una fila
 * en una pestana nueva por visita (ej: "2026-10-22 - Cliente").
 *
 * INSTALACION (una sola vez)
 * 1. Entra a https://script.google.com > Nuevo proyecto.
 * 2. Borra el contenido y pega TODO este archivo.
 * 3. Arriba, elige la funcion "instalar" y pulsa Ejecutar.
 *    - Autoriza los permisos.
 *    - Usa el libro indicado en SPREADSHEET_ID (o crea uno nuevo si esta vacio).
 *    - Crea la pestana INICIO y genera el TOKEN.
 *    - El token queda en el registro de ejecucion.
 * 4. Implementar > Nueva implementacion > Tipo: Aplicacion web.
 *    - Ejecutar como: Yo
 *    - Quien tiene acceso: Cualquier persona
 * 5. Copia la URL /exec y el token en config.json (o en el repositorio
 *    si usas el modo remoto irm).
 *
 * Si ya tienes un libro y quieres usarlo, pon su ID en SPREADSHEET_ID
 * antes de ejecutar instalar().
 *
 * NOTA: al editar una implementacion existente la URL se mantiene.
 */

// ============================ CONFIGURACION ============================

// Deja vacio para que instalar() genere uno aleatorio.
// Si pones una clave aqui, esa manda sobre la generada.
var TOKEN = '';

// Libro maestro donde caeran las visitas. Si lo dejas vacio, instalar() crea uno.
var SPREADSHEET_ID = '1WrRVx2yOePh6XUfbYTUiWP3MuncRZqkrWJyiwbux29U';

// Pestana usada cuando el registro no trae nombre de visita.
var HOJA_POR_DEFECTO = 'Visitas';

// ========================================================================

/**
 * POST desde WigglesMini.ps1
 * Body: { token, id, schema, hoja, registro: { ... } }
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return respuesta_({ ok: false, error: 'Sin cuerpo' });
    }

    var datos = JSON.parse(e.postData.contents);

    var esperado = tokenActual();
    if (!esperado) {
      return respuesta_({ ok: false, error: 'Servidor sin token: ejecuta instalar()' });
    }
    if (String(datos.token || '') !== String(esperado)) {
      return respuesta_({ ok: false, error: 'Token invalido' });
    }

    var id = String(datos.id || '');
    if (!id) {
      return respuesta_({ ok: false, error: 'Falta id' });
    }

    var registro = datos.registro || {};
    if (typeof registro !== 'object' || Array.isArray(registro)) {
      return respuesta_({ ok: false, error: 'Registro invalido' });
    }

    var libro = libro_();
    var nombreHoja = nombreHoja_(datos.hoja || datos.cliente || HOJA_POR_DEFECTO);

    var bloqueo = LockService.getScriptLock();
    bloqueo.waitLock(20000);
    try {
      var hoja = libro.getSheetByName(nombreHoja);
      if (!hoja) {
        hoja = libro.insertSheet(nombreHoja);
      }

      var cabeceras = asegurarCabeceras_(hoja, registro);

      // Idempotencia: si el id ya esta en la columna A, no se duplica.
      var repetido = hoja.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
      if (repetido) {
        return respuesta_({ ok: true, duplicado: true, fila: repetido.getRow(), hoja: nombreHoja });
      }

      var fila = [];
      for (var i = 0; i < cabeceras.length; i++) {
        var campo = cabeceras[i];
        if (campo === 'ID') { fila.push(id); continue; }
        if (campo === 'Recibido') { fila.push(ahora_()); continue; }
        fila.push(valorSeguro_(registro[campo]));
      }

      hoja.appendRow(fila);
      var numeroFila = hoja.getLastRow();
      aplicarFormato_(hoja, numeroFila, cabeceras.length);

      return respuesta_({ ok: true, hoja: nombreHoja, fila: numeroFila, columnas: cabeceras.length });
    } finally {
      bloqueo.releaseLock();
    }
  } catch (err) {
    return respuesta_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/**
 * GET:
 *  - Sin parametros: prueba de despliegue.
 *  - ?verificar=ID&token=...: confirma si un registro llego al libro.
 *    Lo usa WigglesCore para no borrar la cola si el envio fallo.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};

  if (p.verificar) {
    var esperado = tokenActual();
    if (!esperado) {
      return respuesta_({ ok: false, error: 'Servidor sin token: ejecuta instalar()' });
    }
    if (String(p.token || '') !== String(esperado)) {
      return respuesta_({ ok: false, error: 'Token invalido' });
    }

    var libro;
    try {
      libro = libro_();
    } catch (err) {
      return respuesta_({ ok: false, error: String(err && err.message ? err.message : err) });
    }

    var hojas = libro.getSheets();
    for (var i = 0; i < hojas.length; i++) {
      var hallado = hojas[i].getRange('A:A')
        .createTextFinder(String(p.verificar))
        .matchEntireCell(true)
        .findNext();
      if (hallado) {
        return respuesta_({
          ok: true,
          encontrado: true,
          hoja: hojas[i].getName(),
          fila: hallado.getRow()
        });
      }
    }
    return respuesta_({ ok: true, encontrado: false });
  }

  var ping = p.ping ? true : false;
  return respuesta_({
    ok: true,
    app: 'Wiggles Mini',
    version: 1,
    hoja_por_defecto: HOJA_POR_DEFECTO,
    ping: ping
  });
}

// ------------------------------ Auxiliares ------------------------------

function respuesta_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function tokenActual() {
  if (TOKEN) { return TOKEN; }
  try {
    return PropertiesService.getScriptProperties().getProperty('TOKEN') || '';
  } catch (err) {
    return '';
  }
}

function libro_() {
  var props = PropertiesService.getScriptProperties();
  var id = (SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID') || '').trim();

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  var activo = SpreadsheetApp.getActiveSpreadsheet();
  if (!activo) {
    throw new Error('Ejecuta instalar() una vez (o configura SPREADSHEET_ID).');
  }
  return activo;
}

/**
 * INSTALADOR. Ejecutar una sola vez desde el editor de Apps Script.
 *
 * - Si el script es independiente, crea el libro "Wiggles - Inventario".
 * - Si esta pegado dentro de un libro, usa ese libro.
 * - Genera el TOKEN y lo guarda en las propiedades del script.
 * - Crea/actualiza la pestana INICIO con el token y las instrucciones.
 *
 * Al terminar, revisa el registro de ejecucion: ahi aparece el token.
 */
function instalar() {
  var props = PropertiesService.getScriptProperties();
  var id = (SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID') || '').trim();
  var libro;
  var creado = false;

  if (id) {
    libro = SpreadsheetApp.openById(id);
  } else {
    try { libro = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { libro = null; }
    if (!libro) {
      libro = SpreadsheetApp.create('Wiggles - Inventario');
      creado = true;
    }
    id = libro.getId();
  }

  props.setProperty('SPREADSHEET_ID', id);

  var token = tokenActual();
  if (!token) {
    token = generarToken_();
    props.setProperty('TOKEN', token);
  } else {
    props.setProperty('TOKEN', token);
  }

  var url = '';
  try {
    url = ScriptApp.getService().getUrl() || '';
  } catch (err) {
    url = '';
  }

  var hoja = libro.getSheetByName('INICIO');
  if (!hoja) {
    var hojas = libro.getSheets();
    if (hojas.length === 1 && hojas[0].getLastRow() === 0) {
      hoja = hojas[0].setName('INICIO');
    } else {
      hoja = libro.insertSheet('INICIO', 0);
    }
  }

  hoja.clear();
  hoja.setColumnWidth(1, 240);
  hoja.setColumnWidth(2, 620);
  hoja.setHiddenGridlines(true);

  var filas = [
    ['WIGGLES MINI - LIBRO DE INVENTARIO', ''],
    ['', ''],
    ['Token', token],
    ['Webhook (/exec)', url || '(todavia no implementado)'],
    ['Libro', libro.getUrl()],
    ['Pestana por defecto', HOJA_POR_DEFECTO],
    ['', ''],
    ['COMO TERMINAR LA INSTALACION', ''],
    ['1', 'Implementar > Nueva implementacion > Aplicacion web.'],
    ['2', 'Ejecutar como: Yo. Quien tiene acceso: Cualquier persona.'],
    ['3', 'Copia la URL /exec y el token en config.json (o en el repositorio).'],
    ['4', 'Haz una visita de prueba y revisa que aparezca una pestana nueva.'],
    ['', ''],
    ['IMPORTANTE', ''],
    ['', 'No compartas este libro con clientes: contiene el token.'],
    ['', 'Cada visita crea su propia pestana automaticamente.'],
    ['', 'Si necesitas otro token, ejecuta rotarToken().']
  ];

  hoja.getRange(1, 1, filas.length, 2).setValues(filas);
  hoja.getRange(1, 1, 1, 2).merge()
    .setFontSize(15).setFontWeight('bold')
    .setBackground('#1F6F43').setFontColor('#FFFFFF')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 38);

  hoja.getRange(8, 1, 1, 2).merge()
    .setFontWeight('bold').setBackground('#E8F5E9').setFontColor('#1F6F43');
  hoja.getRange(14, 1, 1, 2).merge()
    .setFontWeight('bold').setBackground('#FFF3E0').setFontColor('#B26A00');

  hoja.getRange(3, 1, 5, 1).setFontWeight('bold');
  hoja.getRange(3, 2).setFontFamily('Consolas').setFontSize(11);
  hoja.getRange(4, 2).setFontFamily('Consolas').setFontSize(10);
  hoja.getRange(5, 2).setFontFamily('Consolas').setFontSize(10);

  var mensaje =
    'WIGGLES MINI\n\n' +
    (creado ? 'Libro creado: ' + libro.getUrl() + '\n\n' : 'Libro: ' + libro.getUrl() + '\n\n') +
    'TOKEN:\n' + token + '\n\n' +
    'Webhook: ' + (url || 'implementa la app web para obtenerlo') + '\n\n' +
    'El token tambien quedo en la pestana INICIO.';

  Logger.log(mensaje);
  try {
    SpreadsheetApp.getUi().alert('Wiggles Mini', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    // En proyectos independientes no hay interfaz: el mensaje queda en el log.
  }

  return { libro: libro.getUrl(), token: token, webhook: url };
}

/**
 * Genera un token nuevo y lo guarda. Util si el anterior se filtro.
 */
function rotarToken() {
  var token = generarToken_();
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  var libro = libro_();
  var hoja = libro.getSheetByName('INICIO');
  if (hoja) {
    hoja.getRange(3, 1, 1, 2).setValues([['Token', token]]);
    hoja.getRange(3, 2).setFontFamily('Consolas').setFontSize(11);
  }
  Logger.log('Token nuevo: ' + token + '\nActualiza config.json / repositorio.');
  return token;
}

/**
 * Muestra el token actual en el registro de ejecucion.
 */
function mostrarToken() {
  var token = tokenActual();
  Logger.log(token ? 'Token actual: ' + token : 'No hay token. Ejecuta instalar().');
  return token;
}

/**
 * Resumen del estado para soporte.
 */
function estado() {
  var props = PropertiesService.getScriptProperties();
  var id = SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID');
  var resumen = {
    libro: id,
    token_configurado: Boolean(tokenActual()),
    webhook: '',
    pestanas: []
  };
  try { resumen.webhook = ScriptApp.getService().getUrl() || ''; } catch (err) {}
  if (id) {
    try {
      var libro = SpreadsheetApp.openById(id);
      resumen.pestanas = libro.getSheets().map(function (h) { return h.getName(); });
    } catch (err) {}
  }
  Logger.log(JSON.stringify(resumen, null, 2));
  return resumen;
}

function generarToken_() {
  return Utilities.getUuid().replace(/-/g, '').toLowerCase();
}

function ahora_() {
  var libro = libro_();
  return Utilities.formatDate(new Date(), libro.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Limpia el nombre de la pestana: quita caracteres invalidos y limita a 95.
 */
function nombreHoja_(texto) {
  var limpio = String(texto || '')
    .replace(/[\[\]\:\*\?\/\\]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  if (!limpio) { limpio = HOJA_POR_DEFECTO; }
  if (limpio.length > 95) { limpio = limpio.substring(0, 95).trim(); }
  return limpio;
}

/**
 * Devuelve las cabeceras reales de la hoja. Si la hoja esta vacia las crea;
 * si el registro trae campos nuevos los agrega a la derecha.
 */
function asegurarCabeceras_(hoja, registro) {
  var campos = Object.keys(registro);
  var deseadas = ['ID', 'Recibido'].concat(campos);

  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, deseadas.length).setValues([deseadas]);
    formatearCabecera_(hoja, deseadas.length);
    return deseadas;
  }

  var ultimaCol = Math.max(hoja.getLastColumn(), 1);
  var actuales = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0]
    .map(function (v) { return String(v).trim(); })
    .filter(function (v) { return v !== ''; });

  var faltantes = deseadas.filter(function (c) { return actuales.indexOf(c) === -1; });

  if (faltantes.length > 0) {
    var inicio = actuales.length + 1;
    hoja.getRange(1, inicio, 1, faltantes.length).setValues([faltantes]);
    formatearCabecera_(hoja, inicio + faltantes.length - 1);
    actuales = actuales.concat(faltantes);
  }

  return actuales;
}

function formatearCabecera_(hoja, columnas) {
  var rango = hoja.getRange(1, 1, 1, columnas);
  rango
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F6F43')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setWrap(false);

  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 30);

  if (hoja.getLastRow() <= 1) {
    for (var c = 1; c <= columnas; c++) {
      hoja.setColumnWidth(c, c === 1 ? 40 : (c === 2 ? 140 : 150));
    }
  }
}

function aplicarFormato_(hoja, fila, columnas) {
  if (fila === 1) { return; }

  hoja.getRange(fila, 1, 1, columnas)
    .setVerticalAlignment('top')
    .setWrap(false);

  if (fila % 2 === 0) {
    hoja.getRange(fila, 1, 1, columnas).setBackground('#F3F7F4');
  }

  // Filtro que crece con la hoja.
  var filtro = hoja.getFilter();
  if (!filtro) {
    hoja.getRange(1, 1, Math.max(fila, 2), columnas).createFilter();
  }
}

/**
 * Evita que un valor de texto se interprete como formula.
 */
function valorSeguro_(valor) {
  if (valor === null || valor === undefined) { return ''; }
  if (typeof valor === 'number' || typeof valor === 'boolean') { return valor; }

  var texto = String(valor);
  if (texto.length > 45000) { texto = texto.substring(0, 45000); }
  if (/^[=+@]/.test(texto)) { texto = "'" + texto; }
  return texto;
}

/**
 * Prueba manual desde el editor: Ejecutar > probar_envio.
 */
function probar_envio() {
  var libro = libro_();
  var hoja = libro.insertSheet('PRUEBA ' + Utilities.formatDate(new Date(), libro.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmmss'));
  hoja.getRange(1, 1, 1, 3).setValues([['ID', 'Recibido', 'Cliente']]);
  Logger.log('Hoja de prueba creada: ' + hoja.getName());
}
