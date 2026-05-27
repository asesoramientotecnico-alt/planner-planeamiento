/**
 * 📋 PLANNER COLABORATIVO – Backend en Google Apps Script
 *
 * Este script convierte tu Google Sheet en una API web que la app HTML puede usar
 * para leer y modificar tareas en tiempo real.
 *
 * ▶ INSTALACIÓN (UNA SOLA VEZ):
 *   1. En tu Google Sheet: Extensiones › Apps Script
 *   2. Borrá todo el contenido del editor y pegá ESTE archivo entero
 *   3. Guardá (💾 o Ctrl+S)
 *   4. Andá al menú "Planner" que aparecerá en tu Sheet y clic en
 *      "🚀 Instalar el planner" – te guía paso a paso
 *
 * ▶ ACTUALIZACIÓN A LA VERSIÓN CON FECHAS (Paso 2):
 *   Después de pegar esta versión, andá al menú
 *   "📋 Planner › 🛠 Migrar columnas (Fecha Inicio / Creación)" UNA vez.
 *   Eso agrega los encabezados de las columnas 9 y 10 sin tocar tus datos.
 *
 * No tenés que tocar nada más en este archivo.
 */

const SHEET_NAME = 'Tareas';
const HEADER_ROW = 3;
const DATA_START_ROW = 4;
const COLUMNS = {
  id: 1, area: 2, tarea: 3, responsable: 4,
  fecha: 5, estado: 6, prioridad: 7, notas: 8,
  fechaInicio: 9, fechaCreacion: 10
};
const NUM_COLS = 10;

// ═══════════════════════════════════════════════════════════════════
// MENÚ EN EL SHEET (se ve al abrir el archivo)
// ═══════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Planner')
    .addItem('🚀 Instalar / Reinstalar el planner', 'mostrarInstalador')
    .addItem('🛠 Migrar columnas (Fecha Inicio / Creación)', 'migrarColumnas')
    .addItem('🔗 Ver mi URL del planner', 'mostrarURL')
    .addItem('❓ Ayuda', 'mostrarAyuda')
    .addToUi();
}

function mostrarInstalador() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: -apple-system, sans-serif; padding: 20px; color: #1a1a1a; }
      h2 { color: #1F3864; margin-top: 0; }
      ol li { margin: 10px 0; line-height: 1.5; }
      .btn { background: #1F3864; color: white; padding: 12px 24px; border: none;
             border-radius: 6px; font-size: 14px; cursor: pointer; font-weight: bold; }
      .btn:hover { background: #2a4d8f; }
      .note { background: #FFF8DC; padding: 10px; border-left: 4px solid #FFD966;
              margin: 15px 0; font-size: 13px; }
      code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
    </style>
    <h2>🚀 Vamos a instalar tu planner</h2>
    <p>Hacé clic en el botón. Google te va a pedir permiso para que este script pueda leer y modificar el Sheet. Es normal — es <b>tu</b> script funcionando sobre <b>tu</b> Sheet.</p>
    <ol>
      <li>Clic en <b>Implementar planner</b> abajo.</li>
      <li>Google va a abrir una ventana para autorizar. Elegí tu cuenta.</li>
      <li>Va a decir "Google no verificó esta aplicación" — clic en <b>Configuración avanzada</b> → <b>Ir a (nombre) (no seguro)</b>. Es seguro porque es tu propio script.</li>
      <li>Aceptá los permisos.</li>
      <li>Vas a ver una <b>URL</b> que aparece automáticamente — copiala y pegala en el HTML del planner (donde dice <code>GOOGLE_SCRIPT_URL</code>).</li>
    </ol>
    <div class="note">⚠ Si ya hiciste esto antes y querés generar una URL nueva, también funciona — la URL vieja deja de andar.</div>
    <button class="btn" onclick="google.script.run.withSuccessHandler(mostrarResultado).implementar()">Implementar planner</button>
    <div id="resultado" style="margin-top:20px;"></div>
    <script>
      function mostrarResultado(url) {
        if (url) {
          document.getElementById('resultado').innerHTML =
            '<div style="background:#E8F5E9;padding:15px;border-left:4px solid #4CAF50;border-radius:4px;">' +
            '<b>✅ ¡Listo!</b><br><br>Esta es tu URL. Copiala (Ctrl+C) y pegala en el HTML del planner:<br><br>' +
            '<input type="text" value="' + url + '" style="width:100%;padding:8px;font-family:monospace;font-size:12px;" readonly onclick="this.select();" />' +
            '</div>';
        }
      }
    </script>
  `).setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 Instalación del planner');
}

function implementar() {
  // Asegurarse de que la hoja Tareas exista y tenga la estructura correcta
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('No se encuentra la hoja "Tareas". Asegurate de estar en el Planner_Seguimiento.xlsx.');
  }

  // Guardamos el ID del spreadsheet en propiedades para que doGet/doPost lo usen
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  // Devolvemos la URL del web app actual (si ya existe) o le decimos al usuario que la genere
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) return url;
  } catch (e) {}

  // Si no hay deployment, le mostramos las instrucciones para crearlo
  return null;
}

/**
 * Agrega los encabezados de las columnas nuevas (Fecha Inicio = 9, Fecha Creación = 10)
 * sin tocar los datos existentes. Es seguro correrlo varias veces.
 */
function migrarColumnas() {
  const sheet = getSheet();
  const ui = SpreadsheetApp.getUi();
  sheet.getRange(HEADER_ROW, COLUMNS.fechaInicio).setValue('Fecha Inicio');
  sheet.getRange(HEADER_ROW, COLUMNS.fechaCreacion).setValue('Fecha Creación');
  ui.alert('✅ Listo',
    'Se agregaron los encabezados "Fecha Inicio" (columna I) y "Fecha Creación" (columna J).\n\n' +
    'Tus tareas existentes quedan sin esas fechas (es normal). Cuando edites o crees una tarea, ' +
    'se completan solas. La fecha de creación se setea automáticamente al crear y no se modifica después.',
    ui.ButtonSet.OK);
}

function mostrarURL() {
  let url;
  try {
    url = ScriptApp.getService().getUrl();
  } catch (e) {}
  const html = url
    ? `<div style="font-family:-apple-system,sans-serif;padding:20px;">
         <h3 style="color:#1F3864;">🔗 Tu URL del planner</h3>
         <p>Pegá esta URL en el HTML donde dice <code>GOOGLE_SCRIPT_URL</code>:</p>
         <input type="text" value="${url}" style="width:100%;padding:8px;font-family:monospace;font-size:12px;" readonly onclick="this.select();" />
       </div>`
    : `<div style="font-family:-apple-system,sans-serif;padding:20px;">
         <h3 style="color:#C00000;">⚠ Todavía no implementaste el planner</h3>
         <p>Andá a <b>📋 Planner › 🚀 Instalar el planner</b> primero.</p>
       </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(550).setHeight(220),
    'URL del planner'
  );
}

function mostrarAyuda() {
  const html = `
    <div style="font-family:-apple-system,sans-serif;padding:20px;line-height:1.6;">
      <h3 style="color:#1F3864;">❓ Ayuda</h3>
      <p><b>¿Cómo funciona?</b><br>
      Este script convierte tu Sheet en una mini API. La app HTML del planner usa esa API
      para leer y modificar tareas. Los cambios se ven en el Sheet en tiempo real, y al revés también.</p>
      <p><b>¿Es seguro?</b><br>
      Sí. El script solo puede acceder a este Sheet. Cualquier persona con la URL puede leer y
      modificar las tareas, así que <b>no compartas la URL públicamente</b> — solo con tu equipo.</p>
      <p><b>¿Y si quiero más privacidad?</b><br>
      Al implementar, podés elegir "Solo usuarios de mi organización" en vez de "Cualquiera con el enlace"
      (solo si todos tienen Google Workspace de la misma empresa).</p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(550).setHeight(380),
    'Ayuda'
  );
}

// ═══════════════════════════════════════════════════════════════════
// API: doGet (lectura) y doPost (escritura)
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) {
      return jsonResponse({ ok: true, tareas: [] });
    }
    const values = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, NUM_COLS).getValues();
    const tareas = values
      .map((row, idx) => ({
        rowIndex: DATA_START_ROW + idx,
        id: row[0],
        area: row[1],
        tarea: row[2],
        responsable: row[3],
        fecha: fmtDate(row[4]),
        estado: row[5],
        prioridad: row[6],
        notas: row[7],
        fechaInicio: fmtDate(row[8]),
        fechaCreacion: fmtDate(row[9])
      }))
      .filter(t => t.tarea); // solo filas con descripción
    return jsonResponse({ ok: true, tareas: tareas, timestamp: new Date().toISOString() });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const sheet = getSheet();

    if (action === 'create') {
      const lastRow = Math.max(sheet.getLastRow(), HEADER_ROW);
      const newRow = lastRow + 1;
      const newId = getNextId(sheet);
      const t = body.tarea;
      const now = new Date();
      sheet.getRange(newRow, 1, 1, NUM_COLS).setValues([[
        newId,
        t.area || '',
        t.tarea || '',
        t.responsable || '',
        t.fecha ? new Date(t.fecha) : '',
        t.estado || 'Pendiente',
        t.prioridad || 'Media',
        t.notas || '',
        t.fechaInicio ? new Date(t.fechaInicio) : '',
        now  // Fecha Creación: automática, no editable
      ]]);
      if (t.fecha)       sheet.getRange(newRow, COLUMNS.fecha).setNumberFormat('dd/mm/yyyy');
      if (t.fechaInicio) sheet.getRange(newRow, COLUMNS.fechaInicio).setNumberFormat('dd/mm/yyyy');
      sheet.getRange(newRow, COLUMNS.fechaCreacion).setNumberFormat('dd/mm/yyyy');
      return jsonResponse({ ok: true, id: newId, rowIndex: newRow });
    }

    if (action === 'update') {
      const row = body.rowIndex;
      if (!row || row < DATA_START_ROW) throw new Error('rowIndex inválido');
      const t = body.tarea;
      // Columnas 2..8 (área..notas)
      sheet.getRange(row, 2, 1, 7).setValues([[
        t.area || '', t.tarea || '', t.responsable || '',
        t.fecha ? new Date(t.fecha) : '',
        t.estado || 'Pendiente', t.prioridad || 'Media', t.notas || ''
      ]]);
      if (t.fecha) sheet.getRange(row, COLUMNS.fecha).setNumberFormat('dd/mm/yyyy');
      // Columna 9 (Fecha Inicio). NO tocamos la 10 (Fecha Creación) para preservarla.
      sheet.getRange(row, COLUMNS.fechaInicio).setValue(t.fechaInicio ? new Date(t.fechaInicio) : '');
      if (t.fechaInicio) sheet.getRange(row, COLUMNS.fechaInicio).setNumberFormat('dd/mm/yyyy');
      return jsonResponse({ ok: true });
    }

    if (action === 'delete') {
      const row = body.rowIndex;
      if (!row || row < DATA_START_ROW) throw new Error('rowIndex inválido');
      sheet.deleteRow(row);
      return jsonResponse({ ok: true });
    }

    if (action === 'updateField') {
      // Actualización rápida de un solo campo (para drag&drop, cambio de estado)
      const row = body.rowIndex;
      const colName = body.field;
      if (!COLUMNS[colName]) throw new Error('campo inválido: ' + colName);
      if (colName === 'fechaCreacion') throw new Error('la fecha de creación no se puede modificar');
      sheet.getRange(row, COLUMNS[colName]).setValue(body.value);
      return jsonResponse({ ok: true });
    }

    throw new Error('Acción desconocida: ' + action);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function getSheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Hoja "Tareas" no encontrada');
  return sheet;
}

function getNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 1;
  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues().flat();
  const max = Math.max(0, ...ids.filter(x => typeof x === 'number'));
  return max + 1;
}

/** Formatea una celda de fecha a 'yyyy-MM-dd'. Si no es fecha, devuelve string o ''. */
function fmtDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v ? String(v) : '';
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
