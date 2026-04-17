/**
 * ================================================================
 * SyncToFirebase.gs — IBIME Portal de Calificaciones
 * Sincroniza Google Sheets → Firebase Realtime Database
 *
 * SETUP: Las credenciales se leen desde PropertiesService.
 *        Ejecuta setup_guardarSecrets() en Code.gs antes de usar.
 * ================================================================
 */

var SYNC_CONFIG = {
  HOJA_ALUMNOS: "BASE ALUMNOS",
  HOJA_MAESTRO: "Maestro",

  // Índices de columnas — BASE ALUMNOS
  BD_MATRICULA : 0,
  BD_NOMBRE    : 1,
  BD_TUTOR     : 2,
  BD_GPO_ESP   : 3,
  BD_GPO_ING   : 4,
  BD_CORREO    : 5,

  // Índices de columnas — Maestro
  MAE_PROFESOR : 0,
  MAE_GRUPO    : 1,
  MAE_ALUMNO   : 2,
  MAE_CORREO   : 3,
  MAE_FECHA    : 4,
  MAE_ACTIVIDAD: 5,
  MAE_CALIF    : 6,
  MAE_MATERIA  : 7
};

// ── Entrada principal ────────────────────────────────────────────
function syncAll() {
  Logger.log("⏳ Iniciando sincronización completa...");
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var alumnos        = _syncLeerAlumnos(ss);
  var calificaciones = _syncLeerCalificaciones(ss);

  _syncEscribirFirebase("/alumnos",        alumnos);
  _syncEscribirFirebase("/calificaciones", calificaciones);
  _syncEscribirFirebase("/meta/ultimaSync", new Date().toISOString());

  Logger.log("✅ Sincronización completa.");
}

// ── Crear trigger horario (ejecutar UNA SOLA VEZ) ────────────────
function crearTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("syncAll").timeBased().everyHours(1).create();
  Logger.log("✅ Trigger creado: syncAll cada hora.");
}

// ── Leer alumnos ─────────────────────────────────────────────────
function _syncLeerAlumnos(ss) {
  var hoja  = ss.getSheetByName(SYNC_CONFIG.HOJA_ALUMNOS);
  var data  = hoja.getDataRange().getValues();
  var obj   = {};

  for (var i = 1; i < data.length; i++) {
    var fila      = data[i];
    var matricula = String(fila[SYNC_CONFIG.BD_MATRICULA]).trim();
    if (!matricula) continue;

    obj[_sanitizarClave(matricula)] = {
      matricula: matricula,
      nombre   : String(fila[SYNC_CONFIG.BD_NOMBRE]  || "").trim(),
      tutor    : String(fila[SYNC_CONFIG.BD_TUTOR]    || "No asignado").trim(),
      grupoEsp : String(fila[SYNC_CONFIG.BD_GPO_ESP]  || "---").trim(),
      grupoIng : String(fila[SYNC_CONFIG.BD_GPO_ING]  || "---").trim(),
      correo   : String(fila[SYNC_CONFIG.BD_CORREO]   || "").toLowerCase().trim()
    };
  }
  return obj;
}

// ── Leer calificaciones ──────────────────────────────────────────
function _syncLeerCalificaciones(ss) {
  var hoja  = ss.getSheetByName(SYNC_CONFIG.HOJA_MAESTRO);
  var data  = hoja.getDataRange().getValues();
  var calif = {};

  for (var j = 1; j < data.length; j++) {
    var fila   = data[j];
    var correo = String(fila[SYNC_CONFIG.MAE_CORREO] || "").toLowerCase().trim();
    if (!correo) continue;

    var clave = _sanitizarClave(correo);
    if (!calif[clave]) calif[clave] = {};

    var fechaStr = "---";
    if (fila[SYNC_CONFIG.MAE_FECHA] instanceof Date) {
      fechaStr = Utilities.formatDate(fila[SYNC_CONFIG.MAE_FECHA], "GMT-6", "dd/MM/yyyy");
    } else {
      fechaStr = String(fila[SYNC_CONFIG.MAE_FECHA] || "---");
    }

    var rawCalif = fila[SYNC_CONFIG.MAE_CALIF];
    var calNum   = null;
    if (rawCalif !== "" && rawCalif !== null) {
      var p = parseFloat(String(rawCalif).replace(",", "."));
      if (!isNaN(p)) calNum = p;
    }

    var idx = Object.keys(calif[clave]).length;
    calif[clave][idx] = {
      profesor   : String(fila[SYNC_CONFIG.MAE_PROFESOR]  || "Sin Profesor").trim(),
      grupo      : String(fila[SYNC_CONFIG.MAE_GRUPO]     || "N/A").trim(),
      alumno     : String(fila[SYNC_CONFIG.MAE_ALUMNO]    || "").trim(),
      correo     : correo,
      fecha      : fechaStr,
      actividad  : String(fila[SYNC_CONFIG.MAE_ACTIVIDAD] || "Actividad").trim(),
      calificacion: calNum,
      materia    : String(fila[SYNC_CONFIG.MAE_MATERIA]   || "General").trim()
    };
  }
  return calif;
}

// ── Escribir en Firebase via REST ────────────────────────────────
function _syncEscribirFirebase(path, data) {
  var config = _getFirebaseConfig();  // definida en Code.gs
  if (!config.url || !config.secret) {
    Logger.log("❌ Firebase no configurado.");
    return;
  }

  var resp = UrlFetchApp.fetch(
    config.url + path + ".json?auth=" + config.secret,
    { method: "PUT", contentType: "application/json",
      payload: JSON.stringify(data), muteHttpExceptions: true }
  );

  var code = resp.getResponseCode();
  if (code !== 200)
    Logger.log("❌ Error en " + path + " — HTTP " + code + ": " + resp.getContentText());
  else
    Logger.log("✅ Escrito: " + path);
}

// ── Sanitizar claves para Firebase ──────────────────────────────
function _sanitizarClave(str) {
  return str.replace(/[.#$\[\]\/]/g, "_");
}
