/**
 * 📋 PLANNER COLABORATIVO – Backend en Google Apps Script
 *
 * ▶ INSTALACIÓN (UNA SOLA VEZ):
 *   1. Extensiones › Apps Script → borrá todo → pegá este archivo → Guardá
 *   2. Menú "📋 Planner" › "🚀 Instalar / Reinstalar el planner"
 *
 * ▶ ACTUALIZACIÓN A ESTA VERSIÓN (Paso 2+3):
 *   Después de pegar, andá a:
 *   "📋 Planner › 🛠 Migrar columnas (Fecha Inicio / Creación)"  — agrega cols 9 y 10
 *   "📋 Planner › ⚙ Crear hoja Config"                          — crea la hoja de configuración
 *   Luego reimplementá (misma URL, solo versión nueva).
 */

const SHEET_NAME        = 'Tareas';
const CONFIG_SHEET_NAME = 'Config';
const HEADER_ROW        = 3;
const DATA_START_ROW    = 4;
const COLUMNS = {
  id: 1, area: 2, tarea: 3, responsable: 4,
  fecha: 5, estado: 6, prioridad: 7, notas: 8,
  fechaInicio: 9, fechaCreacion: 10
};
const NUM_COLS = 10;

const DEFAULT_CONFIG = {
  areas:       ['General', 'Negocio', 'PxD', 'OT'],
  colaboradores: [],
  kanbanCols: [
    { status: 'Pendiente',  label: 'Pendiente',  visible: true },
    { status: 'En curso',   label: 'En curso',   visible: true },
    { status: 'Bloqueada',  label: 'Bloqueada',  visible: true },
    { status: 'Completada', label: 'Completada', visible: true },
    { status: 'Cancelada',  label: 'Cancelada',  visible: true }
  ]
};

// ═══════════════════════════════════════════════════════
// MENÚ
// ═══════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Planner')
    .addItem('🚀 Instalar / Reinstalar el planner',           'mostrarInstalador')
    .addItem('🛠 Migrar columnas (Fecha Inicio / Creación)',  'migrarColumnas')
    .addItem('⚙ Crear hoja Config',                          'crearHojaConfig')
    .addItem('🔗 Ver mi URL del planner',                     'mostrarURL')
    .addItem('❓ Ayuda',                                       'mostrarAyuda')
    .addToUi();
}

function mostrarInstalador() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:-apple-system,sans-serif;padding:20px;color:#1a1a1a}
      h2{color:#1F3864;margin-top:0}
      ol li{margin:10px 0;line-height:1.5}
      .btn{background:#1F3864;color:white;padding:12px 24px;border:none;border-radius:6px;font-size:14px;cursor:pointer;font-weight:bold}
      .btn:hover{background:#2a4d8f}
      .note{background:#FFF8DC;padding:10px;border-left:4px solid #FFD966;margin:15px 0;font-size:13px}
      code{background:#f0f0f0;padding:2px 6px;border-radius:3px;font-size:12px}
    </style>
    <h2>🚀 Vamos a instalar tu planner</h2>
    <ol>
      <li>Clic en <b>Implementar planner</b> abajo.</li>
      <li>Google pedirá autorización — elegí tu cuenta.</li>
      <li>"Google no verificó esta app" → <b>Configuración avanzada</b> → <b>Ir a (nombre) (no seguro)</b>.</li>
      <li>Aceptá los permisos.</li>
      <li>Copiá la URL que aparece y pegala en el HTML donde dice <code>GOOGLE_SCRIPT_URL</code>.</li>
    </ol>
    <div class="note">⚠ Si ya lo hiciste antes, también funciona — la URL vieja deja de andar.</div>
    <button class="btn" onclick="google.script.run.withSuccessHandler(mostrarResultado).implementar()">Implementar planner</button>
    <div id="resultado" style="margin-top:20px"></div>
    <script>
      function mostrarResultado(url){
        if(url) document.getElementById('resultado').innerHTML=
          '<div style="background:#E8F5E9;padding:15px;border-left:4px solid #4CAF50;border-radius:4px">'+
          '<b>✅ ¡Listo!</b><br><br>Tu URL:<br><br>'+
          '<input type="text" value="'+url+'" style="width:100%;padding:8px;font-family:monospace;font-size:12px" readonly onclick="this.select()"></div>';
      }
    </script>
  `).setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 Instalación del planner');
}

function implementar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_NAME)) throw new Error('No se encuentra la hoja "Tareas".');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  try { const url = ScriptApp.getService().getUrl(); if (url) return url; } catch(e) {}
  return null;
}

function migrarColumnas() {
  const sheet = getSheet();
  const ui = SpreadsheetApp.getUi();
  sheet.getRange(HEADER_ROW, COLUMNS.fechaInicio  ).setValue('Fecha Inicio');
  sheet.getRange(HEADER_ROW, COLUMNS.fechaCreacion).setValue('Fecha Creación');
  ui.alert('✅ Listo',
    'Se agregaron los encabezados "Fecha Inicio" (col I) y "Fecha Creación" (col J).\n\n' +
    'Tus tareas existentes quedan sin esas fechas — se completan al crear/editar desde ahora.',
    ui.ButtonSet.OK);
}

/**
 * Crea la hoja "Config" con los valores por defecto.
 * Seguro de correr varias veces (no pisa valores existentes).
 */
function crearHojaConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg) {
    cfg = ss.insertSheet(CONFIG_SHEET_NAME);
    cfg.getRange(1,1).setValue('clave');
    cfg.getRange(1,2).setValue('valor_json');
    cfg.setColumnWidth(1, 150);
    cfg.setColumnWidth(2, 500);
  }
  // Escribir defaults solo si la clave no existe
  const keysToSet = {
    areas:         JSON.stringify(DEFAULT_CONFIG.areas),
    colaboradores: JSON.stringify(DEFAULT_CONFIG.colaboradores),
    kanbanCols:    JSON.stringify(DEFAULT_CONFIG.kanbanCols)
  };
  const last = cfg.getLastRow();
  const existing = last > 1
    ? cfg.getRange(2, 1, last-1, 1).getValues().flat().map(String)
    : [];
  Object.entries(keysToSet).forEach(([k, v]) => {
    if (!existing.includes(k)) {
      const newRow = cfg.getLastRow() + 1;
      cfg.getRange(newRow, 1).setValue(k);
      cfg.getRange(newRow, 2).setValue(v);
    }
  });
  SpreadsheetApp.getUi().alert('✅ Hoja Config lista',
    'La hoja "Config" ya existe con los valores por defecto.\n' +
    'Desde la app podés personalizar áreas, colaboradores y columnas del Kanban.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function mostrarURL() {
  let url; try { url = ScriptApp.getService().getUrl(); } catch(e) {}
  const html = url
    ? `<div style="font-family:-apple-system,sans-serif;padding:20px">
         <h3 style="color:#1F3864">🔗 Tu URL</h3>
         <input type="text" value="${url}" style="width:100%;padding:8px;font-family:monospace;font-size:12px" readonly onclick="this.select()">
       </div>`
    : `<div style="font-family:-apple-system,sans-serif;padding:20px">
         <h3 style="color:#C00000">⚠ Todavía no implementaste el planner</h3>
         <p>Andá a <b>📋 Planner › 🚀 Instalar el planner</b> primero.</p>
       </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(550).setHeight(200), 'URL del planner');
}

function mostrarAyuda() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`
      <div style="font-family:-apple-system,sans-serif;padding:20px;line-height:1.6">
        <h3 style="color:#1F3864">❓ Ayuda</h3>
        <p><b>¿Cómo funciona?</b> Este script convierte tu Sheet en una mini API.
        La app HTML lee y escribe tareas en tiempo real.</p>
        <p><b>¿Es seguro?</b> Solo accede a este Sheet.
        No compartas la URL públicamente — solo con tu equipo.</p>
        <p><b>Hoja Config</b> guarda áreas, colaboradores y columnas del Kanban.
        Todos los del equipo ven los mismos cambios.</p>
      </div>`).setWidth(520).setHeight(300), 'Ayuda');
}

// ═══════════════════════════════════════════════════════
// API: doGet  (lectura de tareas + config)
// ═══════════════════════════════════════════════════════
function doGet(e) {
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const ss   = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Hoja "Tareas" no encontrada');

    const lastRow = sheet.getLastRow();
    const tareas  = lastRow < DATA_START_ROW ? [] :
      sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, NUM_COLS)
        .getValues()
        .map((row, idx) => ({
          rowIndex:      DATA_START_ROW + idx,
          id:            row[0],
          area:          row[1],
          tarea:         row[2],
          responsable:   row[3],
          fecha:         fmtDate(row[4]),
          estado:        row[5],
          prioridad:     row[6],
          notas:         row[7],
          fechaInicio:   fmtDate(row[8]),
          fechaCreacion: fmtDate(row[9])
        }))
        .filter(t => t.tarea);

    const config = getConfigData(ss);

    return jsonResponse({ ok: true, tareas, config, timestamp: new Date().toISOString() });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════
// API: doPost (escritura)
// ═══════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    // ── TAREAS ──────────────────────────────────────
    if (action === 'create') {
      const sheet   = getSheet();
      const newRow  = Math.max(sheet.getLastRow(), HEADER_ROW) + 1;
      const newId   = getNextId(sheet);
      const t       = body.tarea;
      const now     = new Date();
      sheet.getRange(newRow, 1, 1, NUM_COLS).setValues([[
        newId, t.area||'', t.tarea||'', t.responsable||'',
        t.fecha      ? new Date(t.fecha)      : '',
        t.estado     || 'Pendiente',
        t.prioridad  || 'Media',
        t.notas      || '',
        t.fechaInicio ? new Date(t.fechaInicio) : '',
        now
      ]]);
      if (t.fecha)       sheet.getRange(newRow, COLUMNS.fecha      ).setNumberFormat('dd/mm/yyyy');
      if (t.fechaInicio) sheet.getRange(newRow, COLUMNS.fechaInicio).setNumberFormat('dd/mm/yyyy');
      sheet.getRange(newRow, COLUMNS.fechaCreacion).setNumberFormat('dd/mm/yyyy');
      return jsonResponse({ ok: true, id: newId, rowIndex: newRow });
    }

    if (action === 'update') {
      const sheet = getSheet();
      const row   = body.rowIndex;
      if (!row || row < DATA_START_ROW) throw new Error('rowIndex inválido');
      const t = body.tarea;
      sheet.getRange(row, 2, 1, 7).setValues([[
        t.area||'', t.tarea||'', t.responsable||'',
        t.fecha ? new Date(t.fecha) : '',
        t.estado||'Pendiente', t.prioridad||'Media', t.notas||''
      ]]);
      if (t.fecha) sheet.getRange(row, COLUMNS.fecha).setNumberFormat('dd/mm/yyyy');
      sheet.getRange(row, COLUMNS.fechaInicio).setValue(t.fechaInicio ? new Date(t.fechaInicio) : '');
      if (t.fechaInicio) sheet.getRange(row, COLUMNS.fechaInicio).setNumberFormat('dd/mm/yyyy');
      // NO tocamos col 10 (fechaCreacion) — es inmutable
      return jsonResponse({ ok: true });
    }

    if (action === 'delete') {
      const sheet = getSheet();
      const row   = body.rowIndex;
      if (!row || row < DATA_START_ROW) throw new Error('rowIndex inválido');
      sheet.deleteRow(row);
      return jsonResponse({ ok: true });
    }

    if (action === 'updateField') {
      const sheet   = getSheet();
      const colName = body.field;
      if (!COLUMNS[colName]) throw new Error('campo inválido: ' + colName);
      if (colName === 'fechaCreacion') throw new Error('fechaCreacion es inmutable');
      sheet.getRange(body.rowIndex, COLUMNS[colName]).setValue(body.value);
      return jsonResponse({ ok: true });
    }

    // ── CONFIG ───────────────────────────────────────
    if (action === 'saveConfig') {
      const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      const ss   = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
      setConfigKey(ss, body.key, body.value);
      return jsonResponse({ ok: true });
    }

    throw new Error('Acción desconocida: ' + action);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════
// CONFIG HELPERS
// ═══════════════════════════════════════════════════════
function getConfigData(ss) {
  const cfg = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg || cfg.getLastRow() < 2) return DEFAULT_CONFIG;
  const rows = cfg.getRange(2, 1, cfg.getLastRow()-1, 2).getValues();
  const result = Object.assign({}, DEFAULT_CONFIG);
  rows.forEach(([k, v]) => {
    if (!k) return;
    try { result[String(k)] = JSON.parse(String(v)); } catch(e) {}
  });
  return result;
}

function setConfigKey(ss, key, value) {
  let cfg = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg) {
    cfg = ss.insertSheet(CONFIG_SHEET_NAME);
    cfg.getRange(1,1).setValue('clave');
    cfg.getRange(1,2).setValue('valor_json');
  }
  const last = cfg.getLastRow();
  if (last >= 2) {
    const keys = cfg.getRange(2,1,last-1,1).getValues().flat().map(String);
    const idx  = keys.indexOf(String(key));
    if (idx >= 0) {
      cfg.getRange(idx+2, 2).setValue(JSON.stringify(value));
      return;
    }
  }
  const newRow = cfg.getLastRow() + 1;
  cfg.getRange(newRow, 1).setValue(key);
  cfg.getRange(newRow, 2).setValue(JSON.stringify(value));
}

// ═══════════════════════════════════════════════════════
// SHEET / ID HELPERS
// ═══════════════════════════════════════════════════════
function getSheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss   = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Hoja "Tareas" no encontrada');
  return sheet;
}

function getNextId(sheet) {
  const last = sheet.getLastRow();
  if (last < DATA_START_ROW) return 1;
  const ids = sheet.getRange(DATA_START_ROW, 1, last - DATA_START_ROW + 1, 1).getValues().flat();
  return Math.max(0, ...ids.filter(x => typeof x === 'number')) + 1;
}

function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v ? String(v) : '';
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
