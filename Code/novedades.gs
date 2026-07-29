// ═══════════════════════════════════════════════════════════════
// Novedades.gs — Centro de Conocimiento · Novedades
// Mensajes cortos por módulo SAP con imagen opcional y vigencia
// definida o permanente. Las alertas de vencimiento son VISUALES
// dentro del panel (colores por urgencia): NO se envían correos.
// Una novedad vencida deja de mostrarse (listarNovedades la filtra por
// fecha) y un disparador diario la borra de verdad de la hoja
// (purgarNovedadesVencidas, ver al final cómo activarlo).
//
// Reutiliza de Centro_conocimiento.gs (mismo proyecto):
//   abrirArchivoB_(), getCarpetaRaizDrive_(), verificarPermiso_()
//   y registrarLog_().
//
// Pestaña "Novedades" del Archivo B (se crea sola si falta):
// A ID | B Módulo | C Novedad | D Imagen | E Publicada |
// F Expira (vacío = permanente) | G Comunicador | H Fuente |
// I CreadoPor | J (reservada, sin uso)
// ═══════════════════════════════════════════════════════════════

var HOJA_NOVEDADES    = "Novedades";
var CARPETA_NOVEDADES = "Novedades";          // subcarpeta de imágenes en Drive

// ════════════════════════════════════════════════════════════
// LECTURA — novedades activas para el panel
// ════════════════════════════════════════════════════════════

// Devuelve las novedades vigentes. Las vencidas no se muestran (se filtran
// aquí por fecha) y el disparador diario purgarNovedadesVencidas() las
// elimina de la hoja.
function listarNovedades() {
  var sheet   = asegurarHojaNovedades_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var tz    = abrirArchivoB_().getSpreadsheetTimeZone();
  var ahora = Date.now();
  var datos = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var items = [];
  for (var i = 0; i < datos.length; i++) {
    var id = (datos[i][0] || "").toString().trim();
    if (!id) continue;                        // ignora filas sin ID
    var expira   = datos[i][5];
    var expiraMs = (expira instanceof Date) ? expira.getTime() : 0;
    if (expiraMs && expiraMs <= ahora) continue;   // ya venció

    var pub = datos[i][4];
    items.push({
      fila:        i + 2,
      id:          id,
      modulo:      (datos[i][1] || "General").toString().trim() || "General",
      texto:       (datos[i][2] || "").toString(),
      imagenId:    (datos[i][3] || "").toString().trim(),
      publicada:   (pub instanceof Date)
                     ? Utilities.formatDate(pub, tz, "dd/MM/yyyy")
                     : (pub || "").toString(),
      expiraMs:    expiraMs,                  // 0 = permanente
      comunicador: (datos[i][6] || "").toString().trim(),
      fuente:      (datos[i][7] || "").toString().trim(),
      creadoPor:   (datos[i][8] || "").toString().trim()
    });
  }
  return items;
}

// ════════════════════════════════════════════════════════════
// CREAR / ELIMINAR (requieren permiso de Admins)
// ════════════════════════════════════════════════════════════

// Crea una novedad. d = { modulo, texto, comunicador, fuente,
//   dias (0 o vacío = permanente), base64?, mime?, nombreArchivo? }.
// Devuelve { ok:true } o { ok:false, error }.
function crearNovedad(d) {
  try {
    var correo = verificarPermiso_();
    if (!d || !(d.texto || "").toString().trim()) {
      throw new Error("Escribe el texto de la novedad");
    }

    // Imagen opcional → subcarpeta "Novedades" en Drive
    var imagenId = "";
    if (d.base64) {
      var contenido = d.base64.indexOf(",") !== -1 ? d.base64.split(",")[1] : d.base64;
      var nombre    = (d.nombreArchivo || "novedad").replace(/[\\\/:*?"<>|]/g, "_");
      var blob = Utilities.newBlob(
        Utilities.base64Decode(contenido),
        d.mime || "application/octet-stream",
        nombre
      );
      var file = getCarpetaNovedadesDrive_().createFile(blob);
      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {
        Logger.log("No se pudo ajustar el uso compartido: " + e.message);
      }
      imagenId = file.getId();
    }

    var dias   = parseFloat(d.dias) || 0;
    var expira = dias > 0 ? new Date(Date.now() + dias * 24 * 60 * 60 * 1000) : "";

    asegurarHojaNovedades_().appendRow([
      Utilities.getUuid(),
      (d.modulo || "General").toString().trim() || "General",
      d.texto.toString().trim(),
      imagenId,
      new Date(),
      expira,                                 // vacío = permanente
      (d.comunicador || "").toString().trim(),
      (d.fuente || "").toString().trim(),
      correo,
      ""                                      // Avisos (interno)
    ]);

    registrarLog_(correo, "CREAR NOVEDAD", resumenNovedad_(d.texto));
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Elimina una novedad por su ID (también manda su imagen a la papelera).
// Como la fecha de expiración queda en el log del sheet, se distingue
// si alguien la quitó antes de tiempo.
function eliminarNovedad(id) {
  try {
    var correo = verificarPermiso_();
    id = (id || "").toString().trim();

    var sheet   = asegurarHojaNovedades_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("La novedad ya no existe. Actualiza el panel.");

    var datos = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    for (var i = 0; i < datos.length; i++) {
      if ((datos[i][0] || "").toString().trim() === id) {
        var texto = (datos[i][2] || "").toString();
        var img   = (datos[i][3] || "").toString().trim();
        sheet.deleteRow(i + 2);
        if (img) {
          try { DriveApp.getFileById(img).setTrashed(true); } catch(e) {}
        }
        registrarLog_(correo, "ELIMINAR NOVEDAD", resumenNovedad_(texto));
        return { ok: true };
      }
    }
    throw new Error("La novedad ya no existe. Actualiza el panel.");
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
// AUXILIARES
// ════════════════════════════════════════════════════════════

// Primeros 80 caracteres del texto (para los logs).
function resumenNovedad_(texto) {
  texto = (texto || "").toString().replace(/\s+/g, " ").trim();
  return texto.length > 80 ? texto.slice(0, 80) + "…" : texto;
}

// Crea la pestaña Novedades si no existe (la columna J queda reservada, sin uso).
function asegurarHojaNovedades_() {
  return asegurarHoja_(HOJA_NOVEDADES,
    ["ID", "Módulo", "Novedad", "Imagen", "Publicada",
     "Expira", "Comunicador", "Fuente", "CreadoPor", "Avisos"],
    [280, 90, 380, 280, 130, 130, 160, 100, 220, 80]);
}

// Subcarpeta "Novedades" dentro de la carpeta raíz de Drive
// (se crea la primera vez que se sube una imagen).
function getCarpetaNovedadesDrive_() {
  var raiz = getCarpetaRaizDrive_();
  var it = raiz.getFoldersByName(CARPETA_NOVEDADES);
  return it.hasNext() ? it.next() : raiz.createFolder(CARPETA_NOVEDADES);
}

// ════════════════════════════════════════════════════════════
// LIMPIEZA AUTOMÁTICA — disparador diario
// ════════════════════════════════════════════════════════════
// Borra DE VERDAD las novedades ya vencidas (y manda su imagen a la
// papelera de Drive), para que la hoja no acumule filas muertas. Las
// novedades PERMANENTES (sin fecha de expiración) nunca se tocan.
// Cada borrado queda anotado en la pestaña Logs como "sistema (limpieza)".
//
// ⚠ CÓMO ACTIVARLO (una sola vez, no pide permisos nuevos):
//   En el editor de Apps Script → menú izquierdo «Activadores» (reloj ⏰)
//   → «Añadir activador» →
//        Función:            purgarNovedadesVencidas
//        Fuente del evento:  Según tiempo
//        Tipo de activador:  Temporizador diario
//        Hora:               p. ej. 3 a. m. – 4 a. m.
//   → Guardar. Desde ahí se ejecuta solo cada día.
function purgarNovedadesVencidas() {
  try {
    var sheet   = asegurarHojaNovedades_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var ahora = Date.now();
    var datos = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

    // Recorremos de ABAJO hacia ARRIBA: al borrar una fila, las de arriba
    // no cambian de número, así que los índices no se descuadran.
    for (var i = datos.length - 1; i >= 0; i--) {
      var expira = datos[i][5];
      if (!(expira instanceof Date)) continue;      // permanente → se conserva
      if (expira.getTime() > ahora)  continue;       // aún vigente

      var img = (datos[i][3] || "").toString().trim();
      if (img) {
        try { DriveApp.getFileById(img).setTrashed(true); } catch(e) {}
      }
      sheet.deleteRow(i + 2);
      registrarLog_("sistema (limpieza)", "PURGAR NOVEDAD", resumenNovedad_(datos[i][2]));
    }
  } catch(e) {
    // Nunca debe fallar de forma ruidosa: es una tarea desatendida.
    Logger.log("purgarNovedadesVencidas falló: " + e.message);
  }
}
