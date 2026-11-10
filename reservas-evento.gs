/**
 * ══════════════════════════════════════════════════════════════════════
 *  KINOMORI — Reservas de la CENA ESPECIAL (Google Apps Script)
 * ══════════════════════════════════════════════════════════════════════
 *
 *  Qué hace:
 *    - doPost  → guarda cada reserva como una fila en una hoja de cálculo
 *    - doGet   → devuelve cuántas plazas hay ocupadas (JSON) para que la
 *                web muestre el aforo real, no un número escrito a mano
 *
 *  INSTALACIÓN (5 minutos, gratis, sin tarjeta):
 *   1. Entra en https://sheets.google.com y crea una hoja nueva.
 *      Ponle nombre: "Kinomori Cena Especial".
 *      Copia el ID de la URL: es lo que va entre /d/ y /edit
 *   2. Entra en https://script.google.com → "Nuevo proyecto"
 *   3. Borra todo y pega ESTE archivo. Guarda (Ctrl+S).
 *   4. Pega el ID de la hoja en SHEET_ID (abajo).
 *   5. En el desplegable de funciones elige "setup" → Ejecutar → da permisos.
 *   6. "Implementar" → "Nueva implementación" → engranaje → "Aplicación web"
 *        · Ejecutar como: Yo
 *        · Quién tiene acceso: Cualquier persona
 *      → Implementar → copia la URL que termina en /exec
 *   7. Pega esa URL en index.html → CONFIG.endpoint
 *
 *  La hoja se crea sola con las cabeceras la primera vez.
 * ══════════════════════════════════════════════════════════════════════
 */

var SHEET_ID   = "";                          // ← pega aquí el ID de tu hoja
var NOMBRE_HOJA = "Reservas";
var PLAZAS_TOTALES = 15;
var EMAIL_AVISO = "carlosgonzvali@gmail.com"; // te llega un email por reserva

var CABECERAS = ["Fecha reserva","Nombre","Apellidos","Personas","Importe MAD",
                 "Pagado","Email","Teléfono","Alergias","Idioma","Evento","Acompañantes"];

/** Crea la hoja con cabeceras si no existe. Ejecuta esto UNA vez. */
function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(CABECERAS);
    sh.getRange(1, 1, 1, CABECERAS.length)
      .setFontWeight("bold").setBackground("#14100c").setFontColor("#f8f1e6");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
  }
  // menú para marcar pagos rápido
  return "Hoja lista: " + ss.getName();
}

/** Cuenta las plazas ya ocupadas (suma de la columna Personas). */
function plazasOcupadas_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(NOMBRE_HOJA);
  if (!sh || sh.getLastRow() < 2) return 0;
  var personas = sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues();
  var total = 0;
  for (var i = 0; i < personas.length; i++) {
    var v = parseInt(personas[i][0], 10);
    if (!isNaN(v)) total += v;
  }
  return total;
}

/** GET → aforo real. Llamada desde la web. */
function doGet(e) {
  var accion = (e && e.parameter && e.parameter.accion) || "";
  var salida = {};
  if (accion === "plazas") {
    var ocupadas = plazasOcupadas_();
    salida = {
      reservadas: ocupadas,
      totales: PLAZAS_TOTALES,
      libres: Math.max(0, PLAZAS_TOTALES - ocupadas)
    };
  } else {
    salida = { ok: true, servicio: "Kinomori reservas" };
  }
  return ContentService.createTextOutput(JSON.stringify(salida))
    .setMimeType(ContentService.MimeType.JSON);
}

/** POST → guarda la reserva. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // evita que dos reservas se pisen
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);

    // control de aforo en servidor (no solo en la web)
    var ocupadas = plazasOcupadas_();
    var piden = parseInt(d.personas, 10) || 1;
    if (ocupadas + piden > PLAZAS_TOTALES) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false, error: "aforo_completo", libres: PLAZAS_TOTALES - ocupadas
      })).setMimeType(ContentService.MimeType.JSON);
    }

    sh.appendRow([
      new Date(),
      d.nombre || "", d.apellidos || "", piden,
      d.importe || "", "NO",
      d.email || "", "'" + (d.telefono || ""),   // apóstrofo: que no lo coma como número
      d.alergias || "", d.idioma || "", d.evento || "cena-especial",
      Array.isArray(d.acompanantes) ? d.acompanantes.join(" / ") : (d.acompanantes || "")
    ]);

    // aviso por email (opcional pero recomendado)
    try {
      MailApp.sendEmail({
        to: EMAIL_AVISO,
        subject: "Reserva Kinomori · " + (d.nombre || "") + " " + (d.apellidos || "")
                 + " · " + piden + " pers.",
        body: "Nueva reserva para la cena especial\n\n"
            + "Nombre: " + d.nombre + " " + d.apellidos + "\n"
            + "Personas: " + piden + "\n"
            + "Importe: " + d.importe + " MAD\n"
            + "Email: " + d.email + "\n"
            + "Teléfono: " + d.telefono + "\n"
            + "Alergias: " + (d.alergias || "ninguna") + "\n\n"
            + "Recuerda marcar PAGADO en la hoja cuando llegue el dinero."
      });
    } catch (err) { /* si falla el email, la reserva ya está guardada */ }

    return ContentService.createTextOutput(JSON.stringify({
      ok: true, libres: Math.max(0, PLAZAS_TOTALES - (ocupadas + piden))
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, error: String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
