/**
 * ================================================================
 * IBIME — Sistema de Importación de Calificaciones
 * Versión GitHub: modular + secrets seguros
 * ================================================================
 *
 * SETUP INICIAL (ejecutar UNA SOLA VEZ desde el editor de GAS):
 *   1. Abre Extensions > Apps Script
 *   2. Corre la función `setup_guardarSecrets()` con tus valores reales
 *   3. Borra esa función o deja que quede comentada
 *
 * DEPLOY:
 *   clasp push  →  Deploy as Web App desde el editor de GAS
 * ================================================================
 */

// ── Constantes globales ──────────────────────────────────────────
var HOJAS = {
  MAESTRO          : "Maestro",
  BASE_MAESTROS    : "BASE MAESTROS",
  BASE_ALUMNOS     : "BASE ALUMNOS",
  GRUPOS           : "GRUPOS",
  ASIGNATURA       : "ASIGNATURA",
  MATRIZ           : "Matriz_Resumen",
  TABLERO          : "TABLERO_AVANCE",
  DASHBOARD        : "DASHBOARD_ENTREGAS",
  BITACORA         : "BITACORA_TIEMPOS"
};

var FECHA_CORTE = "CORTE AL DIA 20/MARZO/2026 (Fecha máxima de carga 23/03/2026)";

// ── Imagen de fondo desde Drive ──────────────────────────────────
var DRIVE_IMG_ID = "1Lwcr0KSKNHzdvmD5yDCEyMeq7iKMlKsW";

// ================================================================
// 1. WEB APP ENTRY POINT
// ================================================================
function doGet() {
  var template = HtmlService.createTemplateFromFile("WebApp");
  template.fechaManual = FECHA_CORTE;
  return template
    .evaluate()
    .setTitle("Importación de calificaciones — Classroom")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getImageUrl() {
  return (
    "data:image/png;base64," +
    Utilities.base64Encode(
      DriveApp.getFileById(DRIVE_IMG_ID).getBlob().getBytes()
    )
  );
}

// ================================================================
// 2. CONFIGURACIÓN DE SECRETS (ejecutar UNA VEZ y luego borrar)
// ================================================================
function setup_guardarSecrets() {
  PropertiesService.getScriptProperties().setProperties({
    FIREBASE_URL   : "https://TU-PROYECTO-default-rtdb.firebaseio.com/",
    FIREBASE_SECRET: "TU_SECRET_AQUI"
  });
  Logger.log("✅ Secrets guardados. Puedes borrar esta función.");
}

function _getFirebaseConfig() {
  var p = PropertiesService.getScriptProperties();
  return {
    url   : p.getProperty("FIREBASE_URL"),
    secret: p.getProperty("FIREBASE_SECRET")
  };
}

// ================================================================
// 3. HELPERS GENERALES
// ================================================================
function capitalizarNombre(nombre) {
  if (!nombre) return "";
  return nombre
    .toLowerCase()
    .split(" ")
    .map(function(p) { return p.charAt(0).toUpperCase() + p.slice(1); })
    .join(" ");
}

function extraerIdDesdeUrl(url) {
  var match = url.match(/\/d\/(.+?)\//);
  return match ? match[1] : null;
}

function obtenerColumnaSimple(nombreHoja, nombreHeader) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) return [];
  var data = hoja.getDataRange().getValues();
  if (data.length < 1) return [];
  var idx = data[0].indexOf(nombreHeader);
  if (idx === -1) return [];
  return data
    .slice(1)
    .map(function(f) { return f[idx]; })
    .filter(function(v) { return v && String(v).trim() !== ""; })
    .map(function(v) { return String(v).trim(); });
}

function obtenerAcronimo(texto) {
  if (!texto) return "N/A";
  var palabras = texto.split(" ").filter(function(w) { return w.length > 2; });
  if (palabras.length === 0) return texto.substring(0, 3).toUpperCase();
  if (palabras.length === 1) return palabras[0].substring(0, 3).toUpperCase();
  return palabras.map(function(w) { return w[0]; }).join("").toUpperCase();
}

// ================================================================
// 4. CONSULTAS DE DATOS (llamadas desde la WebApp)
// ================================================================
function obtenerNombreProfesor(matricula) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJAS.BASE_MAESTROS);
  if (!hoja) return { ok: false, error: "No existe hoja '" + HOJAS.BASE_MAESTROS + "'" };

  var data = hoja.getDataRange().getValues();
  var headers = data[0].map(function(s) { return String(s).trim().toUpperCase(); });
  var idxMat    = headers.indexOf("MATRICULA PROFESOR");
  var idxNombre = headers.indexOf("NOMBRE DEL PROFESOR");

  if (idxMat === -1 || idxNombre === -1)
    return { ok: false, error: "Encabezados incorrectos en " + HOJAS.BASE_MAESTROS };

  matricula = String(matricula).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxMat]).trim() === matricula) {
      return { ok: true, nombre: capitalizarNombre(String(data[i][idxNombre]).trim()) };
    }
  }
  return { ok: false, error: "Matrícula no encontrada" };
}

function obtenerCargaPorMatricula(matricula) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  matricula = String(matricula).trim();

  // 1. Buscar nombre del profesor
  var dataMaestros = ss.getSheetByName(HOJAS.BASE_MAESTROS).getDataRange().getValues();
  var headers = dataMaestros[0].map(function(s) { return String(s).trim().toUpperCase(); });
  var idxMat    = headers.indexOf("MATRICULA PROFESOR");
  var idxNombre = headers.indexOf("NOMBRE DEL PROFESOR");

  if (idxMat === -1 || idxNombre === -1)
    return { ok: false, error: "Encabezados incorrectos en " + HOJAS.BASE_MAESTROS };

  var nombreProfesor = "";
  for (var i = 1; i < dataMaestros.length; i++) {
    if (String(dataMaestros[i][idxMat]).trim() === matricula) {
      nombreProfesor = String(dataMaestros[i][idxNombre]).trim();
      break;
    }
  }
  if (!nombreProfesor) return { ok: false, error: "Matrícula no encontrada." };

  // 2. Obtener grupos para detectar tipo
  var dataGrupos = ss.getSheetByName(HOJAS.GRUPOS).getDataRange().getValues();
  var gruposEsp = [], gruposIng = [];
  for (var g = 1; g < dataGrupos.length; g++) {
    if (dataGrupos[g][0]) gruposEsp.push(String(dataGrupos[g][0]).trim());
    if (dataGrupos[g][1]) gruposIng.push(String(dataGrupos[g][1]).trim());
  }

  // 3. Buscar columna en Matriz_Resumen
  var dataMatriz = ss.getSheetByName(HOJAS.MATRIZ).getDataRange().getValues();
  var encabezados = dataMatriz[0];
  var nombreBusqueda = nombreProfesor.toLowerCase().trim();
  var colIndex = -1;

  for (var c = 0; c < encabezados.length; c++) {
    if (String(encabezados[c]).toLowerCase().trim() === nombreBusqueda) {
      colIndex = c;
      break;
    }
  }

  if (colIndex === -1) {
    return {
      ok    : true,
      nombre: capitalizarNombre(nombreProfesor),
      carga : [],
      aviso : "Profesor encontrado, pero sin asignación en Matriz."
    };
  }

  // 4. Construir carga
  var carga = [];
  var tipoDetectado = "";

  for (var r = 1; r < dataMatriz.length; r++) {
    var cell = dataMatriz[r][colIndex];
    if (!cell || String(cell).trim() === "") continue;

    var grupo   = String(dataMatriz[r][0]).trim();
    var materias = String(cell).split(/\r?\n/);
    var tipo     = gruposIng.includes(grupo) ? "Ingles" : "Español";
    if (!tipoDetectado) tipoDetectado = tipo;

    materias.forEach(function(m) {
      if (m.trim() !== "") carga.push({ grupo: grupo, materia: m.trim(), tipo: tipo });
    });
  }

  return {
    ok             : true,
    nombre         : capitalizarNombre(nombreProfesor),
    carga          : carga,
    tipoPredominante: tipoDetectado || "Español"
  };
}

function obtenerGrupos(tipo) {
  return obtenerColumnaSimple(
    HOJAS.GRUPOS,
    tipo === "Español" ? "GRUPO ESPAÑOL" : "GRUPO INGLES"
  );
}

function obtenerAsignaturas(tipo) {
  return obtenerColumnaSimple(
    HOJAS.ASIGNATURA,
    tipo === "Español" ? "ESPAÑOL" : "INGLES"
  );
}

function validarLink(link) {
  try {
    var id = extraerIdDesdeUrl(link);
    if (!id) throw new Error("URL inválida");
    var ssRemota = SpreadsheetApp.openById(id);
    return { ok: true, nombreHoja: ssRemota.getSheets()[0].getName(), id: id };
  } catch (e) {
    return { ok: false, error: "Sin acceso o link roto" };
  }
}

// ================================================================
// 5. MOTOR PRINCIPAL DE IMPORTACIÓN
// ================================================================
function importarGruposWeb(datos) {
  var tiempoInicio = new Date();
  var lock = LockService.getScriptLock();

  try { lock.waitLock(10000); }
  catch (e) { return { mensajeFinal: "Servidor ocupado. Intenta en unos segundos.", resumen: "" }; }

  try {
    var profesor           = datos.profesor.trim();
    var tipoGrupo          = datos.tipoGrupo;
    var prefijo            = tipoGrupo === "Español" ? "Prof." : "Teacher";
    var profesorConPrefijo = prefijo + " " + profesor;

    var ss           = SpreadsheetApp.getActiveSpreadsheet();
    var hojaDestino  = ss.getSheetByName(HOJAS.MAESTRO);

    if (!hojaDestino) {
      hojaDestino = ss.insertSheet(HOJAS.MAESTRO);
      hojaDestino.appendRow(["Profesor", "Grupo", "Alumno", "Correo", "Fecha", "Actividad", "Calificación", "Asignatura"]);
    }

    // Mapa de duplicados para no importar dos veces el mismo grupo+asignatura
    var mapaDuplicados = {};
    var lastRow = hojaDestino.getLastRow();
    if (lastRow > 1) {
      hojaDestino.getRange(2, 1, lastRow - 1, 8).getValues().forEach(function(fila) {
        mapaDuplicados[fila[0] + "|" + fila[1] + "|" + fila[7]] = true;
      });
    }

    var filasParaAgregar            = [];
    var resumen                     = [];
    var linksProcesadosEnEstaCarga  = {};
    var gruposLog                   = [];
    var asignaturasLog              = [];
    var itemsParaDashboard          = [];

    datos.gruposLinks.forEach(function(item) {
      var grupo      = item.grupo.trim();
      var asignatura = item.asignatura.trim();
      var link       = item.link.trim();

      if (!grupo || !asignatura || !link || linksProcesadosEnEstaCarga[link]) return;
      linksProcesadosEnEstaCarga[link] = true;

      if (mapaDuplicados[profesor + "|" + grupo + "|" + asignatura]) {
        resumen.push("⚠ OMITIDO: " + grupo + " ya existe.");
        return;
      }

      var validacion = validarLink(link);
      if (!validacion.ok) {
        resumen.push("❌ ERROR: Link inválido en " + grupo + ".");
        return;
      }

      try {
        var datosOrigen = SpreadsheetApp.openById(validacion.id)
          .getSheets()[0].getDataRange().getValues();

        if (datosOrigen.length < 5) {
          resumen.push("⚠ OMITIDO: " + grupo + " — archivo con muy pocas filas.");
          return;
        }

        var fechas      = datosOrigen[0];
        var actividades = datosOrigen[1];

        // Detección automática de escala (10 o 100)
        var escala = 10;
        for (var p = 1; p < datosOrigen[2].length; p++) {
          var maxVal = parseFloat(datosOrigen[2][p]);
          if (!isNaN(maxVal) && maxVal > 0) { escala = maxVal; break; }
        }

        for (var i = 3; i < datosOrigen.length; i++) {
          var fila       = datosOrigen[i];
          var filaString = fila.join(" ").toLowerCase();

          if (
            filaString.includes("media de la clase") ||
            filaString.includes("abrir classroom") ||
            filaString.trim() === ""
          ) continue;

          var idxEmail = -1;
          fila.some(function(celda, idx) {
            if (String(celda).includes("@ibime.edu.mx")) { idxEmail = idx; return true; }
          });
          if (idxEmail === -1) continue;

          var correo = String(fila[idxEmail]).trim();
          var nombre = fila.slice(0, idxEmail)
            .filter(function(c) { return c && String(c).trim() !== ""; })
            .join(" ").trim() || "Alumno sin nombre";

          for (var c = idxEmail + 1; c < fila.length; c++) {
            var hF = fechas[c]      ? String(fechas[c]).trim()      : "";
            var hA = actividades[c] ? String(actividades[c]).trim() : "";
            if (!hF && !hA) continue;

            var rawCalif = fila[c];
            var calif    = (rawCalif === "" || rawCalif == null) ? 0 : parseFloat(rawCalif);
            if (escala === 100 && calif > 0) calif = parseFloat((calif / 10).toFixed(1));
            if (calif > 10) calif = 10;

            filasParaAgregar.push([profesor, grupo, nombre, correo, hF, hA, calif, asignatura]);
          }
        }

        // Contar alumnos únicos en este grupo
        var correosUnicos = {};
        filasParaAgregar.forEach(function(f) {
          if (f[1] === grupo) correosUnicos[f[3]] = true;
        });

        gruposLog.push(grupo);
        if (asignaturasLog.indexOf(asignatura) === -1) asignaturasLog.push(asignatura);
        itemsParaDashboard.push(item);

        resumen.push("✅ " + grupo + " — " + Object.keys(correosUnicos).length + " alumno(s).");

      } catch (err) {
        resumen.push("❌ Error en \"" + grupo + "\": " + err.message);
      }
    });

    if (filasParaAgregar.length > 0) {
      hojaDestino.getRange(
        hojaDestino.getLastRow() + 1, 1,
        filasParaAgregar.length, filasParaAgregar[0].length
      ).setValues(filasParaAgregar);

      _enviarAFirebase(filasParaAgregar);
    }

    var duracion = (new Date().getTime() - tiempoInicio.getTime()) / 1000;
    _registrarBitacora(profesor, gruposLog, asignaturasLog, filasParaAgregar.length, duracion);

    if (itemsParaDashboard.length > 0) {
      actualizarTableroAvance(profesor, itemsParaDashboard);
      generarInformeVisualPro(true);
    }

    return {
      mensajeFinal: "¡Listo " + profesorConPrefijo + "! Se importaron " + filasParaAgregar.length + " registros.",
      resumen     : resumen.join("\n")
    };

  } catch (e) {
    return { mensajeFinal: "Error crítico", resumen: e.stack };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// 6. FIREBASE
// ================================================================
function _enviarAFirebase(datos) {
  if (!datos || !datos.length) return;

  var config = _getFirebaseConfig();
  if (!config.url || !config.secret) {
    console.error("❌ Firebase no configurado. Ejecuta setup_guardarSecrets() primero.");
    return;
  }

  var payload   = {};
  var timestamp = new Date().getTime();
  datos.forEach(function(fila, index) {
    payload["reg_" + timestamp + "_" + index] = {
      profesor  : fila[0], grupo     : fila[1],
      alumno    : fila[2], correo    : fila[3],
      fecha_act : fila[4], actividad : fila[5],
      calif     : fila[6], asignatura: fila[7],
      sync      : new Date().toISOString()
    };
  });

  try {
    var resp = UrlFetchApp.fetch(
      config.url + "calificaciones.json?auth=" + config.secret,
      { method: "patch", contentType: "application/json",
        payload: JSON.stringify(payload), muteHttpExceptions: true }
    );
    console.log("Firebase: " + resp.getResponseCode());
  } catch (e) {
    console.error("Error Firebase: " + e.toString());
  }
}

// ================================================================
// 7. BITÁCORA
// ================================================================
function _registrarBitacora(prof, grups, mats, cant, tiem) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hojaLog = ss.getSheetByName(HOJAS.BITACORA) || ss.insertSheet(HOJAS.BITACORA);
    if (hojaLog.getLastRow() === 0)
      hojaLog.appendRow(["Fecha", "Profesor", "Grupos", "Materias", "Filas", "Segundos"]);
    hojaLog.appendRow([new Date(), prof, grups.join(", "), mats.join(", "), cant, tiem]);
  } catch (e) { console.error("Error bitácora: " + e); }
}

// ================================================================
// 8. TABLERO DE AVANCE
// ================================================================
function actualizarTableroAvance(profesor, gruposProcesados) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDash  = ss.getSheetByName(HOJAS.TABLERO);

  if (!hojaDash) {
    hojaDash = ss.insertSheet(HOJAS.TABLERO);
    hojaDash.setTabColor("#FF00FF");
    hojaDash.getRange("A1").setValue("DOCENTE / GRUPOS").setFontWeight("bold");
    hojaDash.setFrozenColumns(1);
    hojaDash.setFrozenRows(1);
    _inicializarTablero(ss, hojaDash);
  } else if (hojaDash.getLastColumn() < 2) {
    _inicializarTablero(ss, hojaDash);
  }

  var norm = function(txt) {
    return String(txt).trim().toUpperCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  };

  var data              = hojaDash.getDataRange().getValues();
  var encabezadosNorm   = data[0].map(norm);
  var nombreProfNorm    = norm(capitalizarNombre(profesor));
  var filaSheets        = -1;

  for (var r = 1; r < data.length; r++) {
    if (norm(data[r][0]) === nombreProfNorm) { filaSheets = r + 1; break; }
  }
  if (filaSheets === -1) {
    hojaDash.appendRow([capitalizarNombre(profesor)]);
    filaSheets = hojaDash.getLastRow();
  }

  gruposProcesados.forEach(function(item) {
    var grupoNorm = norm(item.grupo);
    var acronimo  = obtenerAcronimo(item.asignatura);
    var colSheets = -1;

    for (var c = 1; c < encabezadosNorm.length; c++) {
      if (encabezadosNorm[c] === grupoNorm) { colSheets = c + 1; break; }
    }
    if (colSheets === -1) return;

    var celda       = hojaDash.getRange(filaSheets, colSheets);
    var valorActual = String(celda.getValue()).trim();
    if (!valorActual.includes(acronimo)) {
      celda.setValue(valorActual ? valorActual + ", " + acronimo : acronimo)
           .setBackground("#d4edda");
    }
  });
}

function reconstruirTableroManual() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMaestro = ss.getSheetByName(HOJAS.MAESTRO);
  if (!hojaMaestro) { SpreadsheetApp.getUi().alert("❌ No se encontró la hoja 'Maestro'"); return; }

  var datos = hojaMaestro.getDataRange().getValues();
  if (datos.length < 2) { SpreadsheetApp.getUi().alert("⚠️ La hoja Maestro está vacía."); return; }

  var hojaDash = ss.getSheetByName(HOJAS.TABLERO);
  if (!hojaDash) {
    hojaDash = ss.insertSheet(HOJAS.TABLERO);
    hojaDash.setTabColor("#FF00FF");
    hojaDash.getRange("A1").setValue("DOCENTE / GRUPOS").setFontWeight("bold");
    hojaDash.setFrozenColumns(1); hojaDash.setFrozenRows(1);
    _inicializarTablero(ss, hojaDash);
  }

  var mapa = {};
  for (var i = 1; i < datos.length; i++) {
    var prof  = String(datos[i][0]).trim();
    var grupo = String(datos[i][1]).trim();
    var mat   = String(datos[i][7]).trim();
    if (!prof || !grupo || !mat) continue;
    if (!mapa[prof]) mapa[prof] = {};
    if (!mapa[prof][grupo]) mapa[prof][grupo] = new Set();
    mapa[prof][grupo].add(obtenerAcronimo(mat));
  }

  var total = 0;
  for (var profNombre in mapa) {
    var items = [];
    for (var g in mapa[profNombre]) {
      mapa[profNombre][g].forEach(function(acron) { items.push({ grupo: g, asignatura: acron }); });
    }
    actualizarTableroAvance(profNombre, items);
    total++;
  }
  SpreadsheetApp.getUi().alert("✅ Tablero reconstruido.\n" + total + " profesores procesados.");
}

function _inicializarTablero(ss, hojaDash) {
  var grupos = [].concat(
    obtenerColumnaSimple(HOJAS.GRUPOS, "GRUPO ESPAÑOL"),
    obtenerColumnaSimple(HOJAS.GRUPOS, "GRUPO INGLES")
  ).filter(function(v, i, arr) { return arr.indexOf(v) === i && v; });

  if (grupos.length > 0) {
    hojaDash.getRange(1, 2, 1, grupos.length)
      .setValues([grupos]).setFontWeight("bold").setBackground("#f3f3f3");
  }
}

// ================================================================
// 9. DASHBOARD DE ENTREGAS
// ================================================================
function generarInformeVisualPro(silencioso) {
  silencioso = silencioso || false;
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMatriz  = ss.getSheetByName(HOJAS.MATRIZ);
  var hojaMaestro = ss.getSheetByName(HOJAS.MAESTRO);
  if (!hojaMatriz || !hojaMaestro) return;

  var norm = function(txt) {
    return String(txt).trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  };

  // Carga esperada
  var dataMatriz        = hojaMatriz.getDataRange().getValues();
  var grupos            = dataMatriz.map(function(r) { return String(r[0]).trim(); });
  var profesores        = dataMatriz[0];
  var cargaEsperada     = {};
  var nombresOriginales = {};

  for (var col = 1; col < profesores.length; col++) {
    var nombreProf = String(profesores[col]).trim();
    if (!nombreProf) continue;
    var clave = norm(nombreProf);
    cargaEsperada[clave]     = [];
    nombresOriginales[clave] = capitalizarNombre(nombreProf);

    for (var fila = 1; fila < dataMatriz.length; fila++) {
      var contenido = dataMatriz[fila][col];
      if (!contenido || String(contenido).trim() === "") continue;
      String(contenido).split(/\r?\n/).forEach(function(m) {
        if (m.trim() !== "")
          cargaEsperada[clave].push(norm(grupos[fila]) + " (" + norm(m.trim()) + ")");
      });
    }
  }

  // Entregas reales
  var dataMaestro        = hojaMaestro.getDataRange().getValues();
  var entregasRealizadas = {};
  for (var i = 1; i < dataMaestro.length; i++) {
    var p = norm(dataMaestro[i][0]);
    var g = norm(dataMaestro[i][1]);
    var m = norm(dataMaestro[i][7]);
    if (!p || !g || !m) continue;
    if (!entregasRealizadas[p]) entregasRealizadas[p] = new Set();
    entregasRealizadas[p].add(g + " (" + m + ")");
  }

  // Construir filas
  var filasInforme = [["ESTATUS", "DOCENTE", "ENTREGAS", "% PROGRESO", "GRUPOS/MATERIAS PENDIENTES"]];

  Object.keys(cargaEsperada)
    .sort(function(a, b) { return nombresOriginales[a].localeCompare(nombresOriginales[b]); })
    .forEach(function(clave) {
      var esperados     = cargaEsperada[clave];
      var realizadosSet = entregasRealizadas[clave] || new Set();
      var faltantes     = esperados.filter(function(x) { return !realizadosSet.has(x); });
      var total         = esperados.length;
      var hecho         = total - faltantes.length;
      var porcentaje    = total > 0 ? hecho / total : 0;
      var estatus       = hecho === 0 ? "❌ SIN INICIAR" : faltantes.length > 0 ? "⏳ EN PROCESO" : "✅ COMPLETADO";
      var listaFaltantes = faltantes.length > 0
        ? "• " + faltantes.map(function(f) { return capitalizarNombre(f); }).join("\n• ")
        : "¡Todo entregado!";
      filasInforme.push([estatus, nombresOriginales[clave], hecho + " de " + total, porcentaje, listaFaltantes]);
    });

  // Volcar y estilizar
  var hojaInfo = ss.getSheetByName(HOJAS.DASHBOARD);
  if (hojaInfo) { hojaInfo.clear(); hojaInfo.clearConditionalFormatRules(); }
  else hojaInfo = ss.insertSheet(HOJAS.DASHBOARD);

  hojaInfo.setTabColor("#202124");
  hojaInfo.getRange(1, 1, filasInforme.length, 5).setValues(filasInforme);

  hojaInfo.getRange("A1:E1")
    .setBackground("#1A73E8").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  hojaInfo.setRowHeight(1, 35);
  [140, 260, 110, 130, 420].forEach(function(w, idx) { hojaInfo.setColumnWidth(idx + 1, w); });

  if (filasInforme.length > 1) {
    var nd = filasInforme.length - 1;
    hojaInfo.getRange(2, 1, nd, 5).setVerticalAlignment("middle");
    hojaInfo.getRange(2, 5, nd, 1).setWrap(true);
    hojaInfo.getRange(2, 3, nd, 2).setHorizontalAlignment("center");
    hojaInfo.getRange(2, 4, nd, 1).setNumberFormat("0%");

    var reglas = [
      SpreadsheetApp.newConditionalFormatRule().whenTextContains("✅")
        .setBackground("#D9EAD3").setFontColor("#274E13").setRanges([hojaInfo.getRange(2, 1, nd, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains("⏳")
        .setBackground("#FFF2CC").setFontColor("#BF9000").setRanges([hojaInfo.getRange(2, 1, nd, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains("❌")
        .setBackground("#F4CCCC").setFontColor("#990000").setRanges([hojaInfo.getRange(2, 1, nd, 1)]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue("#EA4335", SpreadsheetApp.InterpolationType.NUMBER, "0")
        .setGradientMidpointWithValue("#FBBC04", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
        .setGradientMaxpointWithValue("#34A853", SpreadsheetApp.InterpolationType.NUMBER, "1")
        .setRanges([hojaInfo.getRange(2, 4, nd, 1)]).build()
    ];
    hojaInfo.setConditionalFormatRules(reglas);
    for (var r = 2; r <= filasInforme.length; r++) {
      var lineas = (filasInforme[r - 1][4].match(/\n/g) || []).length + 1;
      hojaInfo.setRowHeight(r, 21 * Math.max(1, lineas));
    }
  }

  hojaInfo.getRange(1, 1, filasInforme.length, 5)
    .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  hojaInfo.setHiddenGridlines(true);

  if (!silencioso) {
    var comp = filasInforme.slice(1).filter(function(f) { return f[0].includes("✅"); }).length;
    var proc = filasInforme.slice(1).filter(function(f) { return f[0].includes("⏳"); }).length;
    var sin  = filasInforme.slice(1).filter(function(f) { return f[0].includes("❌"); }).length;
    SpreadsheetApp.getUi().alert(
      "✅ Informe generado.\n\n✅ Completados: " + comp +
      "\n⏳ En proceso: " + proc + "\n❌ Sin iniciar: " + sin +
      "\n\nTotal docentes: " + (filasInforme.length - 1)
    );
  } else {
    console.log("Dashboard actualizado automáticamente.");
  }
}

// ================================================================
// 10. UTILIDADES DE PERMISOS
// ================================================================
function forzarPermisos() {
  UrlFetchApp.fetch("https://www.google.com");
}
// ═══════════════════════════════════════════════════════════════════════
// AGREGAR AL FINAL DE Code.gs (después de la línea 694)
// ═══════════════════════════════════════════════════════════════════════

// ================================================================
// 11. MÓDULO DE CONSULTA DE REPORTES — LOGIN & AUTENTICACIÓN
// ================================================================

/**
 * Valida credenciales del profesor: Matrícula + Correo
 * Retorna datos del profesor si es válido
 */
function validarLoginProfesor(matricula, correo) {
  matricula = String(matricula).trim();
  correo = String(correo).toLowerCase().trim();
  
  if (!matricula || !correo) {
    return { ok: false, error: "Matrícula y correo son requeridos" };
  }
  
  // Validar formato de correo
  if (!correo.includes("@ibime.edu.mx")) {
    return { ok: false, error: "Usa un correo institucional @ibime.edu.mx" };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaMAestros = ss.getSheetByName(HOJAS.BASE_MAESTROS);
    if (!hojaMAestros) {
      return { ok: false, error: "Configuración incorrecta: falta " + HOJAS.BASE_MAESTROS };
    }
    
    var data = hojaMAestros.getDataRange().getValues();
    var headers = data[0].map(function(s) { return String(s).trim().toUpperCase(); });
    var idxMat = headers.indexOf("MATRICULA PROFESOR");
    var idxNombre = headers.indexOf("NOMBRE DEL PROFESOR");
    var idxCorreo = headers.indexOf("CORREO") > -1 ? headers.indexOf("CORREO") : -1;
    
    if (idxMat === -1 || idxNombre === -1) {
      return { ok: false, error: "Encabezados incorrectos en BASE MAESTROS" };
    }
    
    // Buscar profesor por matrícula
    for (var i = 1; i < data.length; i++) {
      var matRow = String(data[i][idxMat]).trim();
      if (matRow === matricula) {
        var nombreProf = capitalizarNombre(String(data[i][idxNombre]).trim());
        var correoRow = idxCorreo > -1 ? String(data[i][idxCorreo]).toLowerCase().trim() : "";
        
        // Validar correo (simple o exacto si existe en el sheet)
        if (idxCorreo > -1 && correoRow && correoRow !== correo) {
          return { ok: false, error: "Correo no coincide con la matrícula" };
        }
        
        return {
          ok: true,
          matricula: matricula,
          nombre: nombreProf,
          correo: correo
        };
      }
    }
    
    return { ok: false, error: "Matrícula no encontrada" };
    
  } catch (e) {
    return { ok: false, error: "Error en validación: " + e.toString() };
  }
}

// ================================================================
// 12. OBTENER REPORTE DE PROFESOR (para la consulta)
// ================================================================

/**
 * Obtiene todos los alumnos y sus calificaciones para un profesor
 * Organizado por grupo y ordenado por estado de riesgo
 */
function obtenerReporteProfesor(matricula) {
  matricula = String(matricula).trim();
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    
    if (!hojaMaestro) {
      return { ok: false, error: "No hay datos cargados aún" };
    }
    
    var data = hojaMaestro.getDataRange().getValues();
    var alumnosMap = {};  // { correo: { nombre, grupo, materia, calificaciones: [] } }
    
    // Leer todas las filas del maestro
    for (var i = 1; i < data.length; i++) {
      var fila = data[i];
      var profesor = String(fila[0]).trim();
      var grupo = String(fila[1]).trim();
      var alumno = String(fila[2]).trim();
      var correo = String(fila[3]).toLowerCase().trim();
      var calif = parseFloat(fila[6]) || 0;
      var materia = String(fila[7]).trim();
      
      // Solo incluir registros del profesor que consulta
      if (profesor.toLowerCase() !== _obtenerNombreProfesorPorMatricula(matricula).toLowerCase()) {
        continue;
      }
      
      if (!correo) continue;
      
      if (!alumnosMap[correo]) {
        alumnosMap[correo] = {
          nombre: alumno,
          grupo: grupo,
          materias: {},
          calificaciones: [],
          promedio: 0
        };
      }
      
      alumnosMap[correo].calificaciones.push(calif);
      
      if (!alumnosMap[correo].materias[materia]) {
        alumnosMap[correo].materias[materia] = [];
      }
      alumnosMap[correo].materias[materia].push(calif);
    }
    
    // Calcular promedios y estados
    var alumnos = [];
    for (var correoAl in alumnosMap) {
      var al = alumnosMap[correoAl];
      var promedio = al.calificaciones.length > 0
        ? al.calificaciones.reduce(function(a, b) { return a + b; }, 0) / al.calificaciones.length
        : 0;
      promedio = parseFloat(promedio.toFixed(2));
      
      var estado = _obtenerEstadoCalificacion(promedio);
      
      alumnos.push({
        correo: correoAl,
        nombre: al.nombre,
        grupo: al.grupo,
        promedio: promedio,
        estado: estado,
        materias: al.materias,
        totalCalificaciones: al.calificaciones.length
      });
    }
    
    // Ordenar de peor a mejor (CRÍTICO primero)
    alumnos.sort(function(a, b) {
      var ordenEstatus = { "Crítico": 0, "En Riesgo": 1, "Satisfactorio": 2, "Excelente": 3 };
      return ordenEstatus[a.estado.texto] - ordenEstatus[b.estado.texto];
    });
    
    // Obtener resumen de grupos
    var gruposSet = {};
    for (var ii = 1; ii < data.length; ii++) {
      var prof = String(data[ii][0]).trim();
      var grp = String(data[ii][1]).trim();
      var mat = String(data[ii][7]).trim();
      
      if (prof.toLowerCase() === _obtenerNombreProfesorPorMatricula(matricula).toLowerCase()) {
        if (!gruposSet[grp]) gruposSet[grp] = new Set();
        gruposSet[grp].add(mat);
      }
    }
    
    // Construir resumen de grupos (similar a auditoría)
    var resumenGrupos = [];
    for (var gname in gruposSet) {
      resumenGrupos.push({
        grupo: gname,
        materias: Array.from(gruposSet[gname]),
        status: "✅ Completado"  // Puedes mejorar esto verificando si falta algo
      });
    }
    
    return {
      ok: true,
      alumnos: alumnos,
      gruposResumen: resumenGrupos,
      totalAlumnos: alumnos.length
    };
    
  } catch (e) {
    return { ok: false, error: "Error al obtener reporte: " + e.toString() };
  }
}

// ================================================================
// 13. HELPERS PARA ESTADOS Y CÁLCULOS
// ================================================================

function _obtenerEstadoCalificacion(valor) {
  valor = parseFloat(valor) || 0;
  
  if (valor >= 9) {
    return { emoji: "😎", texto: "Excelente", color: "#3B82F6" };
  }
  if (valor >= 8) {
    return { emoji: "🙂", texto: "Satisfactorio", color: "#10B981" };
  }
  if (valor >= 6) {
    return { emoji: "⚠️", texto: "En Riesgo", color: "#F59E0B" };
  }
  return { emoji: "😓", texto: "Crítico", color: "#EF4444" };
}

function _obtenerNombreProfesorPorMatricula(matricula) {
  var res = obtenerNombreProfesor(matricula);
  return res.ok ? res.nombre : "";
}

// ================================================================
// 14. GUARDAR COMENTARIO DE PROFESOR (Firebase)
// ================================================================

/**
 * Guarda un comentario opcional del profesor sobre un alumno
 * Se almacena en Firebase bajo /comentarios_docentes
 */
function guardarComentarioAlumno(matriculaProf, correoAlumno, comentario) {
  matriculaProf = String(matriculaProf).trim();
  correoAlumno = String(correoAlumno).toLowerCase().trim();
  comentario = String(comentario || "").trim();
  
  if (!comentario) return { ok: false, error: "Comentario vacío" };
  if (comentario.length > 500) return { ok: false, error: "Comentario muy largo (máx 500 caracteres)" };
  
  try {
    var config = _getFirebaseConfig();
    if (!config.url || !config.secret) {
      return { ok: false, error: "Firebase no configurado" };
    }
    
    var payload = {};
    var clave = _sanitizarClave(correoAlumno + "_" + matriculaProf);
    
    payload[clave] = {
      matriculaProfesor: matriculaProf,
      correoAlumno: correoAlumno,
      comentario: comentario,
      timestamp: new Date().toISOString()
    };
    
    var resp = UrlFetchApp.fetch(
      config.url + "comentarios_docentes.json?auth=" + config.secret,
      {
        method: "patch",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    
    if (resp.getResponseCode() === 200) {
      return { ok: true, mensaje: "Comentario guardado" };
    } else {
      return { ok: false, error: "Error al guardar en Firebase" };
    }
    
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

function _sanitizarClave(str) {
  return str.replace(/[.#$\[\]\/]/g, "_");
}
