// ═══════════════════════════════════════════════════════════════
// Code.gs — SAP Tools (Archivo A)
// Proyecto: CENIT Transporte — Seguridad y Accesos SAP
// Contiene: menú, apertura de paneles, registro en Sheets,
//           comparación de nombres y creación de borradores Gmail.
// El Centro de Conocimiento vive en Centro_conocimiento.gs
// (otro archivo de ESTE MISMO proyecto); sus datos, en el Archivo B.
// ═══════════════════════════════════════════════════════════════

// ── Configuración global ─────────────────────────────────────

// Hojas que usan estructura de fila por consultor (sin appendRow).
// Fuente ÚNICA: el panel del Derivador la consume vía getHojasSinConsultor().
var HOJAS_SIN_CONSULTOR = ["SEGA", "SOX Sega", "Basis"];

// PDFs adjuntados automáticamente al crear cada borrador de correo.
// Para agregar un PDF:
//   1. Sube el archivo a Google Drive
//   2. Abre el archivo → copia la URL
//   3. El ID es la parte entre /d/ y /view:
//      drive.google.com/file/d/  ESTE_ES_EL_ID  /view
var PDF_ADJUNTOS = [
  "1jC8FLw9D1IVcUnOpt0Pl59RNw3GwGvtn",
  "1a2QiRr9_NZoDH2KzdOzHDahAtN88ozQV",
];

// ════════════════════════════════════════════════════════════
// MENÚ
// Se crea al abrir la hoja de cálculo.
// ════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙ SAP Tools")
    .addItem("Redactor de Correos Portales",   "abrirRedactor")
    .addItem("Redactor de Correos SAP",        "abrirRedactorSap")
    .addSeparator()
    .addItem("Derivador — ventana pequeña", "abrirTicketsPequeno")
    .addItem("Derivador — ventana grande",  "abrirTicketsGrande")
    .addSeparator()
    .addItem("Centro de Conocimiento",    "abrirRepositorio")
    .addToUi();
}

// ── Redactor de Correos Portales ────────────────────────────
function abrirRedactor() {
  verificarAutorizacion();
  var html = HtmlService.createHtmlOutputFromFile("Redactor_portales")
    .setWidth(1200).setHeight(860);
  SpreadsheetApp.getUi().showModelessDialog(html, "Redactor de Correos Portales");
}

// ── Redactor SAP — Cambio de Clave ──────────────────────────
function abrirRedactorSap() {
  verificarAutorizacion();
  var html = HtmlService.createHtmlOutputFromFile("Redactor_sap")
    .setWidth(1200).setHeight(860);
  SpreadsheetApp.getUi().showModelessDialog(html, "Redactor SAP — Cambio de Clave");
}

// ── Derivador de tickets ─────────────────────────────────────
function abrirTicketsPequeno() { abrirTickets(700,  650); }
function abrirTicketsGrande()  { abrirTickets(1200, 900); }

function abrirTickets(ancho, alto) {
  var html = HtmlService.createHtmlOutputFromFile("Derivador")
    .setWidth(ancho).setHeight(alto);
  SpreadsheetApp.getUi().showModelessDialog(html, "Derivador SAP");
}

// ════════════════════════════════════════════════════════════
// CENTRO DE CONOCIMIENTO
// Todo su código (abrirRepositorio, carpetas, recursos, permisos
// y logs) vive en el archivo Centro_conocimiento.gs de este mismo
// proyecto. Sus datos viven en el Archivo B (ver ID_ARCHIVO_B allí).
// ════════════════════════════════════════════════════════════

// ── Verificar y forzar autorización OAuth ────────────────────
// Toca Gmail y Drive con operaciones de solo lectura para que
// Google solicite los permisos al usuario antes de abrir el panel.
// Así el error de permisos aparece aquí (con la pantalla de autorización)
// y no dentro del diálogo cuando el usuario intenta crear el borrador.
function verificarAutorizacion() {
  try {
    GmailApp.getDrafts();      // fuerza scope gmail.compose
    DriveApp.getRootFolder();  // fuerza scope drive.readonly
  } catch(e) {
    // Si lanza error es porque no hay autorización → Google pedirá permisos
    throw new Error("Se requiere autorización. Vuelve a intentarlo.");
  }
}

// ════════════════════════════════════════════════════════════
// REGISTRO DE TICKETS EN SHEETS
// ════════════════════════════════════════════════════════════

// Obtiene una hoja por nombre; lanza error si no existe.
function getSheet(nombre) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    var disponibles = ss.getSheets().map(function(s) { return s.getName(); });
    throw new Error("Hoja '" + nombre + "' no existe. Disponibles: " + disponibles.join(", "));
  }
  return sheet;
}

// Expone la lista de hojas sin consultor al panel del Derivador.
// Así la lista existe en un solo lugar (la constante de arriba) y no hay
// que mantener una copia duplicada dentro del HTML.
function getHojasSinConsultor() {
  return HOJAS_SIN_CONSULTOR;
}

// Registra un ticket en la hoja indicada.
// Si la hoja es de tipo "sin consultor", busca la fila del consultor;
// si no, agrega una fila nueva al final.
function registrarEnSheet(datos) {
  var nombreHoja = datos.hoja.trim();
  var sheet      = getSheet(nombreHoja);
  if (HOJAS_SIN_CONSULTOR.indexOf(nombreHoja) !== -1) {
    registrarEnFilaConsultor(sheet, datos);
  } else {
    sheet.appendRow([
      datos.asignador, datos.caso, datos.fecha,
      datos.hora, datos.titulo, datos.consultor
    ]);
  }
  return "OK";
}

// Busca la primera fila vacía asignada al consultor y escribe el ticket.
function registrarEnFilaConsultor(sheet, datos) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 2) throw new Error("La hoja no tiene estructura válida");

  // Localizar la columna "consultor"
  var headers      = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colConsultor = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().toLowerCase().indexOf("consultor") !== -1) {
      colConsultor = i + 1; break;
    }
  }
  if (colConsultor === -1) throw new Error("No se encontró la columna 'consultor' en la hoja");

  // Encontrar la fila vacía del consultor buscado
  var allData     = sheet.getRange(2, 1, lastRow - 1, colConsultor).getValues();
  var filaDestino = -1;
  for (var i = 0; i < allData.length; i++) {
    var consultorCelda = allData[i][colConsultor - 1].toString().trim();
    var celdaA         = allData[i][0].toString().trim();
    if (consultorCelda !== "" && celdaA === "" &&
        nombresCoinciden(datos.consultor.trim(), consultorCelda)) {
      filaDestino = i + 2; break;
    }
  }
  if (filaDestino === -1) throw new Error("No hay espacio disponible para: " + datos.consultor);

  sheet.getRange(filaDestino, 1, 1, 5).setValues([[
    datos.asignador, datos.caso, datos.fecha, datos.hora, datos.titulo
  ]]);
}

// ── Comparación de nombres ───────────────────────────────────
// Compara ignorando tildes, puntuación, artículos y orden de palabras.
function nombresCoinciden(n1, n2) {
  var a = normalizar(n1), b = normalizar(n2);
  if (a === b) return true;

  var ignorar = ["DE","DEL","LA","EL","LOS","LAS","Y"];
  function filtrar(t) {
    return t.split(/\s+/).filter(function(p) {
      return p.length > 2 && ignorar.indexOf(p) === -1;
    });
  }

  var p1 = filtrar(a), p2 = filtrar(b);
  if (!p2.length) return false;

  var ok = 0;
  for (var i = 0; i < p2.length; i++)
    for (var j = 0; j < p1.length; j++)
      if (p1[j] === p2[i]) { ok++; break; }

  return ok === p2.length;
}

function normalizar(t) {
  return t.toUpperCase()
    .replace(/[.\-_,]/g, " ")
    .replace(/Á/g,"A").replace(/É/g,"E").replace(/Í/g,"I")
    .replace(/Ó/g,"O").replace(/Ú/g,"U").replace(/Ñ/g,"N")
    .replace(/\s+/g," ").trim();
}

// ════════════════════════════════════════════════════════════
// REDACTOR — CREAR BORRADOR EN GMAIL
// Recibe desde Redactor_portales.html el objeto `d` con los campos:
//   para, cc, caso, titulo, nombreCliente,
//   accion, idSap, claveTemporal, correoCliente
// Devuelve { ok: true, asunto } o { ok: false, error }
// ════════════════════════════════════════════════════════════
function crearBorradorGmail(d) {
  try {
    // Corrige mojibake UTF-8→Latin-1 y reemplaza "contraseña" por "clave"
    function fix(t) {
      if (!t) return "";
      return t
        .replace(/Ã±/g,"ñ").replace(/Ã©/g,"é").replace(/Ã¡/g,"á")
        .replace(/Ã­/g,"í").replace(/Ã³/g,"ó").replace(/Ãº/g,"ú")
        .replace(/contraseña/gi,"clave").replace(/password/gi,"clave");
    }

    // Asunto del borrador
    var asunto = "Re: " + fix(d.caso) + " - " + fix(d.titulo);
    if (d.idSap) asunto += " - " + d.idSap;

    // Clave: naranja si existe, rojo si está pendiente
    var clave = d.claveTemporal
      ? esc(d.claveTemporal)
      : "<span style='color:#e53e3e;font-style:italic;'>⚠ pendiente</span>";

    // Cuerpo del correo en HTML
    var html =
        "<div style='font-family:Arial,sans-serif;font-size:13px;color:#1e293b;line-height:1.7;max-width:680px;'>"
      + "<p style='margin:0 0 14px 0;'>Cordial Saludo,</p>"
      + "<p style='margin:0 0 14px 0;'>Estimado(a) <strong>" + esc(d.nombreCliente) + "</strong>,</p>"
      + "<p style='margin:0 0 14px 0;'>Por medio del presente correo se informa que la solicitud "
        + "<strong>" + esc(d.caso) + "</strong>, referente a <strong>" + fix(esc(d.titulo)) + "</strong>, "
        + "ha sido gestionada satisfactoriamente de acuerdo con el requerimiento registrado. "
        + "A continuación se detalla la gestión realizada:</p>"
      + "<p style='margin:0 0 6px 0;'><strong>Detalle de la gestión:</strong></p>"
      + "<p style='margin:0 0 14px 0;padding:10px 14px;background:#f8fafc;"
        + "border-left:3px solid #0f3460;border-radius:0 6px 6px 0;'>" + esc(d.accion) + "</p>"
      + "<p style='margin:0 0 24px 0;'></p>"
      + "<table style='border-collapse:collapse;margin:0 0 16px 0;'>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "ID / Código SAP:</td><td>" + (d.idSap ? esc(d.idSap) : "—") + "</td></tr>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "Clave temporal:</td><td>" + clave + "</td></tr>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "Correo registrado:</td><td>" + (d.correoCliente ? esc(d.correoCliente) : "—") + "</td></tr>"
      + "</table>"
      + "<div style='background:#fffbeb;border:1px solid #fde68a;border-radius:6px;"
          + "padding:10px 14px;margin:0 0 16px 0;'>"
        + "<strong style='color:#92400e;'>⚠ Importante:</strong> "
        + "<span style='color:#78350f;'>Tiene <strong>24 horas</strong> para cambiar su clave temporal. "
          + "Si no lo hace a tiempo, el acceso expirará y deberá solicitar una nueva clave desde Mati: "
          + "<a href='https://geemidstream-dwp.onbmc.com/dwp/app/#/catalog' "
          + "style='color:#1155cc;font-weight:600;'>Ingresar aquí</a>.<br>"
          + "Su nueva clave debe cumplir estos requisitos: mínimo <strong>30 caracteres</strong>, "
          + "combinar mayúsculas, minúsculas, números y símbolos, y <strong>no incluir</strong> su ID de usuario.</span>"
      + "</div>"
      + "<p style='margin:0;'>Se procede con el cierre del caso <strong>" + esc(d.caso) + "</strong>.</p>"
      + "</div>";

    var opciones = { htmlBody: html };
    if (d.cc && d.cc.trim()) opciones.cc = d.cc.trim();

    // Adjuntar PDFs desde Drive (IDs en PDF_ADJUNTOS al inicio del archivo)
    var adjuntos = [];
    for (var i = 0; i < PDF_ADJUNTOS.length; i++) {
      var id = PDF_ADJUNTOS[i].trim();
      if (!id) continue;
      try {
        adjuntos.push(DriveApp.getFileById(id).getBlob());
      } catch(e) {
        Logger.log("Adjunto no encontrado: " + id + " — " + e.message);
      }
    }
    if (adjuntos.length) opciones.attachments = adjuntos;

    GmailApp.createDraft(d.para, asunto, "", opciones);
    return { ok: true, asunto: asunto };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
// REDACTOR SAP — CREAR BORRADOR EN GMAIL
// Recibe desde Redactor_sap.html el objeto `d` con los campos:
//   para, caso, titulo, nombreCliente,
//   correoCliente, idSap, claveTemporal, sistema, accion
// Sin CC. Devuelve { ok: true, asunto } o { ok: false, error }
// ════════════════════════════════════════════════════════════
function crearBorradorGmailSap(d) {
  try {
    function fix(t) {
      if (!t) return "";
      return t
        .replace(/Ã±/g,"ñ").replace(/Ã©/g,"é").replace(/Ã¡/g,"á")
        .replace(/Ã­/g,"í").replace(/Ã³/g,"ó").replace(/Ãº/g,"ú")
        .replace(/contraseña/gi,"clave").replace(/password/gi,"clave");
    }

    var asunto = "Re: " + fix(d.caso) + " - " + fix(d.titulo);
    if (d.idSap) asunto += " - " + d.idSap;

    var clave = d.claveTemporal
      ? "<span style='font-family:monospace;background:#fffbeb;padding:1px 6px;border-radius:4px;'>"
        + esc(d.claveTemporal) + "</span>"
      : "<span style='color:#e53e3e;font-style:italic;'>⚠ pendiente</span>";

    var html =
        "<div style='font-family:Arial,sans-serif;font-size:13px;color:#1e293b;line-height:1.7;max-width:680px;'>"
      + "<p style='margin:0 0 14px 0;'>Cordial Saludo,</p>"
      + "<p style='margin:0 0 14px 0;'>Estimado(a) <strong>" + esc(d.nombreCliente) + "</strong>,</p>"
      + "<p style='margin:0 0 14px 0;'>Por medio del presente correo se informa que la solicitud "
        + "<strong>" + esc(d.caso) + "</strong>, referente a <strong>" + fix(esc(d.titulo)) + "</strong>, "
        + "ha sido gestionada satisfactoriamente de acuerdo con el requerimiento registrado. "
        + "A continuación se detalla la gestión realizada:</p>"
      + "<p style='margin:0 0 6px 0;'><strong>Detalle de la gestión:</strong></p>"
      + "<p style='margin:0 0 16px 0;padding:10px 14px;background:#f8fafc;"
        + "border-left:3px solid #059669;border-radius:0 6px 6px 0;'>" + esc(d.accion) + "</p>"
      + "<table style='border-collapse:collapse;margin:0 0 16px 0;'>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "ID / Código SAP:</td><td>" + (d.idSap ? esc(d.idSap) : "—") + "</td></tr>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "Clave temporal:</td><td>" + clave + "</td></tr>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "Correo registrado:</td><td>" + (d.correoCliente ? esc(d.correoCliente) : "—") + "</td></tr>"
        + "<tr><td style='padding:4px 18px 4px 0;font-weight:700;color:#0f2a4a;white-space:nowrap;'>"
          + "Sistema SAP:</td><td><strong>" + (d.sistema ? esc(d.sistema) : "—") + "</strong></td></tr>"
      + "</table>"
      + "<div style='background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;"
          + "padding:10px 14px;margin:0 0 16px 0;'>"
        + "<strong style='color:#065f46;'>⚠ Importante:</strong> "
        + "<span style='color:#047857;'>Tiene <strong>24 horas</strong> para actualizar su clave temporal "
          + "en el sistema. Si no lo hace a tiempo, el acceso expirará y deberá solicitar "
          + "un nuevo caso a través de Mati: "
          + "<a href='https://geemidstream-dwp.onbmc.com/dwp/app/#/catalog' "
          + "style='color:#059669;font-weight:600;'>Ingresar aquí</a>.<br>"
          + "Su nueva clave debe combinar <strong>mayúsculas, minúsculas, números y símbolos</strong>, "
          + "y <strong>no incluir</strong> su ID de usuario.</span>"
      + "</div>"
      + "<p style='margin:0;'>Se procede con el cierre del caso <strong>" + esc(d.caso) + "</strong>.</p>"
      + "</div>";

    GmailApp.createDraft(d.para, asunto, "", { htmlBody: html });
    return { ok: true, asunto: asunto };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Escapa caracteres especiales HTML para evitar inyección en el cuerpo del correo
function esc(t) {
  if (!t) return "";
  return t.toString()
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
