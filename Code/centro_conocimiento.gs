// ═══════════════════════════════════════════════════════════════
// Centro_conocimiento.gs — SAP Tools (mismo proyecto del Archivo A)
// Todo el Centro de Conocimiento vive en este archivo:
// panel, carpetas, recursos, permisos por correo y logs.
// Los DATOS viven en otro Google Sheet (Archivo B), al que este
// código llega por ID — nunca por la hoja activa.
//
// Pestañas del Archivo B (se crean solas si faltan):
//   Repositorio → Categoría | Título | Tipo | Enlace / ID Drive |
//                 Descripción | Agregado | ID_Carpeta_Padre
//   Carpetas    → ID | Nombre | ID_Carpeta_Padre | ID_Carpeta_Drive |
//                 CreadoPor | Fecha
//   Logs        → Fecha y hora | Correo | Acción | Elemento
//   Admins      → Correo   (⚠ esta la llenas TÚ: correos autorizados)
// ═══════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN — ⚠ LO ÚNICO QUE DEBES EDITAR ─────────────

// ⚠ EDITAR — ID del Archivo B (el Sheet que guarda los datos).
//   Ábrelo y copia el ID de su URL:
//   docs.google.com/spreadsheets/d/  ESTE_ES_EL_ID  /edit
var ID_ARCHIVO_B = "1AuKABU8F89OXMBiD234ljfzLvMTXfaHA3cgL0gk-H0c";

// ⚠ EDITAR solo si cambias de carpeta — carpeta RAÍZ de Drive donde
//   viven los archivos. Las carpetas creadas desde el panel se crean
//   DENTRO de esta como subcarpetas reales.
//   (Este valor es tu carpeta actual "Repositorio informacion".)
var CARPETA_RAIZ_DRIVE_ID = "1EO-Oqhr6iXuCX-SpSTmiHUTz9MNkFdE-";

// Nombres de las pestañas del Archivo B.
var HOJA_REPOSITORIO = "Repositorio";
var HOJA_CARPETAS    = "Carpetas";
var HOJA_LOGS        = "Logs";
var HOJA_ADMINS      = "Admins";

// ════════════════════════════════════════════════════════════
// ACCESOS BASE
// ════════════════════════════════════════════════════════════

// Abre el Archivo B por su ID. Se cachea por ejecución: cada acción del panel
// lo abre UNA sola vez. (En Apps Script los valores globales se reinician en
// cada llamada, así que nunca queda una referencia vieja entre acciones.)
var _archivoB = null;
function abrirArchivoB_() {
  if (_archivoB) return _archivoB;
  try {
    _archivoB = SpreadsheetApp.openById(ID_ARCHIVO_B);
    return _archivoB;
  } catch(e) {
    throw new Error("No se pudo abrir el Archivo B. Revisa ID_ARCHIVO_B en Centro_conocimiento.gs — " + e.message);
  }
}

// Carpeta raíz de Drive del Centro de Conocimiento.
function getCarpetaRaizDrive_() {
  try {
    return DriveApp.getFolderById(CARPETA_RAIZ_DRIVE_ID);
  } catch(e) {
    throw new Error("No se pudo abrir la carpeta raíz de Drive. Revisa CARPETA_RAIZ_DRIVE_ID — " + e.message);
  }
}

// Devuelve una pestaña del Archivo B, creándola con sus encabezados y estilo
// si no existe. Es el ayudante que usan todos los "asegurarHoja..." de abajo,
// así que crear una pestaña nueva es una sola línea.
//   nombre      → nombre de la pestaña
//   encabezados → títulos de la fila 1 (array)
//   anchos      → ancho de cada columna en el mismo orden (array, opcional)
function asegurarHoja_(nombre, encabezados, anchos) {
  var ss    = abrirArchivoB_();
  var sheet = ss.getSheetByName(nombre);
  if (sheet) return sheet;
  sheet = ss.insertSheet(nombre);
  sheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
    .setFontColor("#ffffff").setFontWeight("bold").setBackground("#0f2a4a");
  if (anchos) {
    for (var i = 0; i < anchos.length; i++) sheet.setColumnWidth(i + 1, anchos[i]);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

// ════════════════════════════════════════════════════════════
// PANEL — apertura y carga de datos
// ════════════════════════════════════════════════════════════

// Abre el panel (opción del menú ⚙ SAP Tools del Archivo A).
function abrirRepositorio() {
  verificarAutorizacion();   // definida en Código.gs (mismo proyecto)
  var html = HtmlService.createHtmlOutputFromFile("Repositorio")
    .setWidth(1500).setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(html, "Centro de Conocimiento — SAP");
}

// Entrega al panel todo lo que necesita en UNA sola llamada:
// carpetas + recursos + URL/gid del Archivo B (para el botón Ficha).
function cargarDatosCC() {
  var ss        = abrirArchivoB_();
  var sheetRepo = asegurarHojaRepositorio_();
  // URL de la carpeta raíz de Drive, para el botón "Abrir Drive" del panel.
  // Va en try/catch: si la carpeta fallara, la carga del panel no se rompe.
  var urlDrive = "";
  try { urlDrive = getCarpetaRaizDrive_().getUrl(); } catch(e) {}
  return {
    carpetas:        leerCarpetas_(),
    recursos:        leerRecursos_(),
    urlArchivoB:     ss.getUrl(),
    gidRepositorio:  sheetRepo.getSheetId(),
    urlCarpetaDrive: urlDrive
  };
}

// ════════════════════════════════════════════════════════════
// PERMISOS — control por correo (pestaña Admins)
// Solo quien esté en Admins puede: agregar recursos, eliminar
// recursos, crear y eliminar carpetas. Ver y navegar es libre.
// ════════════════════════════════════════════════════════════

// Correo del usuario que está usando el panel.
function getCorreoActivo_() {
  var correo = "";
  try { correo = Session.getActiveUser().getEmail(); } catch(e) {}
  if (!correo) {
    try { correo = Session.getEffectiveUser().getEmail(); } catch(e) {}
  }
  return (correo || "").toString().trim().toLowerCase();
}

// ¿El correo está en la pestaña Admins?
function esAdmin_(correo) {
  var sheet   = asegurarHojaAdmins_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;              // lista vacía → nadie autorizado
  var datos = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < datos.length; i++) {
    if ((datos[i][0] || "").toString().trim().toLowerCase() === correo) return true;
  }
  return false;
}

// Valida el permiso ANTES de cualquier acción de escritura.
// Devuelve el correo (se usa para el log) o lanza un error claro
// que el panel muestra tal cual.
function verificarPermiso_() {
  var correo = getCorreoActivo_();
  if (!correo) {
    throw new Error("No se pudo identificar tu cuenta de Google. " +
      "Inicia sesión con tu cuenta corporativa e inténtalo de nuevo.");
  }
  if (!esAdmin_(correo)) {
    throw new Error("Tu cuenta (" + correo + ") no está autorizada para esta acción. " +
      "Pide a un administrador que la agregue en la pestaña Admins.");
  }
  return correo;
}

// Crea la pestaña Admins si no existe (columna A = Correo).
function asegurarHojaAdmins_() {
  return asegurarHoja_(HOJA_ADMINS, ["Correo"], [280]);
}

// ════════════════════════════════════════════════════════════
// LOGS — auditoría (pestaña Logs)
// ════════════════════════════════════════════════════════════

function registrarLog_(correo, accion, elemento) {
  try {
    var sheet = asegurarHojaLogs_();
    var tz    = abrirArchivoB_().getSpreadsheetTimeZone();
    sheet.appendRow([
      Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss"),
      correo, accion, elemento
    ]);
  } catch(e) {
    // El log nunca debe tumbar la acción principal.
    Logger.log("No se pudo escribir el log: " + e.message);
  }
}

function asegurarHojaLogs_() {
  return asegurarHoja_(HOJA_LOGS,
    ["Fecha y hora", "Correo", "Acción", "Elemento"],
    [140, 240, 160, 330]);
}

// ════════════════════════════════════════════════════════════
// CARPETAS — estructura tipo explorador (pestaña Carpetas)
// Cada fila es una carpeta del panel, enlazada a una subcarpeta
// REAL de Drive, anidada según su carpeta padre.
// ID_Carpeta_Padre vacío = carpeta en la raíz.
// ════════════════════════════════════════════════════════════

// Devuelve todas las carpetas como lista de objetos.
function leerCarpetas_() {
  var sheet   = asegurarHojaCarpetas_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var datos = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var items = [];
  for (var i = 0; i < datos.length; i++) {
    var id = (datos[i][0] || "").toString().trim();
    if (!id) continue;                       // ignora filas sin ID
    items.push({
      fila:    i + 2,
      id:      id,
      nombre:  (datos[i][1] || "").toString().trim(),
      idPadre: (datos[i][2] || "").toString().trim(),
      idDrive: (datos[i][3] || "").toString().trim()
    });
  }
  return items;
}

// Carpeta REAL de Drive donde guardar según la carpeta lógica.
// idCarpeta vacío = carpeta raíz.
function getCarpetaDriveDestino_(idCarpeta) {
  idCarpeta = (idCarpeta || "").toString().trim();
  if (!idCarpeta) return getCarpetaRaizDrive_();

  var carpetas = leerCarpetas_();
  for (var i = 0; i < carpetas.length; i++) {
    if (carpetas[i].id === idCarpeta) {
      try {
        return DriveApp.getFolderById(carpetas[i].idDrive);
      } catch(e) {
        throw new Error("La carpeta de Drive asociada a «" + carpetas[i].nombre +
          "» ya no existe o no tienes acceso.");
      }
    }
  }
  throw new Error("La carpeta destino ya no existe. Actualiza el panel.");
}

// Crea una carpeta (requiere permiso). Además de la fila en la hoja,
// crea la subcarpeta REAL en Drive dentro de la carpeta del padre.
// Devuelve { ok:true } o { ok:false, error }.
function crearCarpetaRepositorio(nombre, idPadre) {
  try {
    var correo = verificarPermiso_();

    nombre  = (nombre  || "").toString().trim();
    idPadre = (idPadre || "").toString().trim();
    if (!nombre) throw new Error("Escribe el nombre de la carpeta");

    // Subcarpeta real en Drive, anidada según el padre.
    var carpetaDrive = getCarpetaDriveDestino_(idPadre).createFolder(nombre);

    var sheet = asegurarHojaCarpetas_();
    var tz    = abrirArchivoB_().getSpreadsheetTimeZone();
    sheet.appendRow([
      Utilities.getUuid(),                   // ID interno de la carpeta
      nombre,
      idPadre,                               // vacío = raíz
      carpetaDrive.getId(),                  // carpeta real de Drive
      correo,
      Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm")
    ]);

    registrarLog_(correo, "CREAR CARPETA", nombre);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Elimina una carpeta SOLO si está vacía (sin subcarpetas ni recursos).
// La subcarpeta real de Drive (vacía) se envía a la papelera.
// Devuelve { ok:true } o { ok:false, error }.
function eliminarCarpetaRepositorio(idCarpeta) {
  try {
    var correo = verificarPermiso_();
    idCarpeta  = (idCarpeta || "").toString().trim();

    var carpetas = leerCarpetas_();
    var objetivo = null;
    for (var i = 0; i < carpetas.length; i++) {
      if (carpetas[i].id === idCarpeta) { objetivo = carpetas[i]; break; }
    }
    if (!objetivo) throw new Error("La carpeta ya no existe. Actualiza el panel.");

    // Seguridad: no borrar si tiene subcarpetas…
    for (var i = 0; i < carpetas.length; i++) {
      if (carpetas[i].idPadre === idCarpeta) {
        throw new Error("La carpeta «" + objetivo.nombre + "» tiene subcarpetas. Vacíala primero.");
      }
    }
    // …ni archivos dentro.
    var recursos = leerRecursos_();
    for (var i = 0; i < recursos.length; i++) {
      if (recursos[i].idCarpeta === idCarpeta) {
        throw new Error("La carpeta «" + objetivo.nombre + "» tiene archivos. Vacíala primero.");
      }
    }

    asegurarHojaCarpetas_().deleteRow(objetivo.fila);

    // Enviar a la papelera la subcarpeta real (está vacía).
    try { DriveApp.getFolderById(objetivo.idDrive).setTrashed(true); }
    catch(e) { Logger.log("No se pudo enviar la carpeta de Drive a la papelera: " + e.message); }

    registrarLog_(correo, "ELIMINAR CARPETA", objetivo.nombre);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Crea la pestaña Carpetas si no existe.
function asegurarHojaCarpetas_() {
  return asegurarHoja_(HOJA_CARPETAS,
    ["ID", "Nombre", "ID_Carpeta_Padre", "ID_Carpeta_Drive", "CreadoPor", "Fecha"],
    [280, 200, 280, 280, 220, 130]);
}

// ════════════════════════════════════════════════════════════
// RECURSOS — archivos y enlaces (pestaña Repositorio)
// ID_Carpeta_Padre vacío = recurso en la raíz. Los recursos
// antiguos (sin ese dato) siguen apareciendo en la raíz.
// ════════════════════════════════════════════════════════════

// Devuelve los recursos como lista de objetos. Cada item incluye
// `fila` (número de fila real) para eliminarlo con seguridad.
function leerRecursos_() {
  var sheet   = asegurarHojaRepositorio_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var tz    = abrirArchivoB_().getSpreadsheetTimeZone();
  var datos = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var items = [];
  for (var i = 0; i < datos.length; i++) {
    var titulo = (datos[i][1] || "").toString().trim();
    if (!titulo) continue;                   // ignora filas sin título
    var fecha = datos[i][5];
    items.push({
      fila:        i + 2,
      categoria:   (datos[i][0] || "").toString().trim() || "Sin categoría",
      titulo:      titulo,
      tipo:        (datos[i][2] || "").toString().trim(),
      enlace:      (datos[i][3] || "").toString().trim(),
      descripcion: (datos[i][4] || "").toString().trim(),
      fecha:       (fecha instanceof Date)
                     ? Utilities.formatDate(fecha, tz, "dd/MM/yyyy")
                     : (fecha || "").toString().trim(),
      idCarpeta:   (datos[i][6] || "").toString().trim()
    });
  }
  return items;
}

// Registra un recurso (requiere permiso). Dos modos según el panel:
//   - Archivo: d.base64 + d.nombreArchivo + d.mime → se sube a la
//     subcarpeta de Drive de d.idCarpeta (o a la raíz) y se guarda su ID.
//   - Enlace:  d.enlace → se guarda tal cual (URL o ID de Drive).
// Devuelve { ok:true } o { ok:false, error }.
function registrarRecursoRepositorio(d) {
  try {
    var correo = verificarPermiso_();        // permiso ANTES de subir nada

    if (!d || !d.titulo) throw new Error("Falta el título");
    var enlace = (d.enlace || "").toString().trim();

    if (d.base64) {
      // El dataURL viene como "data:<mime>;base64,<contenido>"
      var contenido = d.base64.indexOf(",") !== -1 ? d.base64.split(",")[1] : d.base64;
      var nombre    = (d.nombreArchivo || d.titulo).replace(/[\\\/:*?"<>|]/g, "_");
      var blob = Utilities.newBlob(
        Utilities.base64Decode(contenido),
        d.mime || "application/octet-stream",
        nombre
      );
      var file = getCarpetaDriveDestino_(d.idCarpeta).createFile(blob);

      // Compartir con el dominio para que la vista previa funcione.
      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {
        Logger.log("No se pudo ajustar el uso compartido: " + e.message);
      }
      enlace = file.getId();
    }

    if (!enlace) throw new Error("Falta el archivo o el enlace");

    var sheet = asegurarHojaRepositorio_();
    var tz    = abrirArchivoB_().getSpreadsheetTimeZone();
    sheet.appendRow([
      d.categoria || "Sin categoría",
      d.titulo,
      d.tipo || "Documento",
      enlace,
      d.descripcion || "",
      Utilities.formatDate(new Date(), tz, "dd/MM/yyyy"),
      (d.idCarpeta || "").toString().trim()
    ]);

    registrarLog_(correo, "AGREGAR RECURSO", d.titulo);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Elimina la fila de un recurso (requiere permiso). Antes de tocar nada
// verifica que el título siga en esa fila (si otro usuario borró filas,
// los números cambian). Si borrarArchivo es true y el enlace es de
// Drive, envía el archivo a la papelera.
// Devuelve { ok:true } o { ok:false, error }.
function eliminarRecursoRepositorio(fila, titulo, borrarArchivo) {
  try {
    var correo = verificarPermiso_();

    var sheet = asegurarHojaRepositorio_();
    if (fila < 2 || fila > sheet.getLastRow()) {
      throw new Error("Fila fuera de rango. Actualiza el panel.");
    }

    var datosFila  = sheet.getRange(fila, 1, 1, 7).getValues()[0];
    var tituloHoja = (datosFila[1] || "").toString().trim();
    if (tituloHoja !== (titulo || "").trim()) {
      throw new Error("La hoja cambió desde la última carga. Actualiza el panel e inténtalo de nuevo.");
    }

    if (borrarArchivo) {
      var id = extraerIdDrive_((datosFila[3] || "").toString());
      if (id) {
        try { DriveApp.getFileById(id).setTrashed(true); }
        catch(e) { Logger.log("No se pudo enviar a la papelera: " + e.message); }
      }
    }

    sheet.deleteRow(fila);
    registrarLog_(correo, "ELIMINAR RECURSO", tituloHoja);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Crea la pestaña Repositorio si no existe (7 columnas). Si existe de
// una versión anterior, completa los encabezados que falten (F y G):
// así los recursos antiguos migran sin perderse — su columna G queda
// vacía y aparecen en la raíz del panel.
function asegurarHojaRepositorio_() {
  var ss    = abrirArchivoB_();
  var sheet = ss.getSheetByName(HOJA_REPOSITORIO);
  var estilo = function(rango, texto) {
    rango.setValue(texto)
      .setFontColor("#ffffff").setFontWeight("bold").setBackground("#0f2a4a");
  };

  if (sheet) {
    if (!sheet.getRange(1, 6).getValue()) estilo(sheet.getRange(1, 6), "Agregado");
    if (!sheet.getRange(1, 7).getValue()) estilo(sheet.getRange(1, 7), "ID_Carpeta_Padre");
    return sheet;
  }

  return asegurarHoja_(HOJA_REPOSITORIO,
    ["Categoría", "Título", "Tipo", "Enlace / ID Drive",
     "Descripción", "Agregado", "ID_Carpeta_Padre"],
    [160, 250, 90, 330, 330, 90, 280]);
}

// Extrae el ID de Drive de un enlace, o lo devuelve si ya es un ID suelto.
function extraerIdDrive_(enlace) {
  if (!enlace) return "";
  var m = enlace.match(/\/d\/([-\w]{20,})/) || enlace.match(/[?&]id=([-\w]{20,})/);
  if (m) return m[1];
  return /^[-\w]{20,}$/.test(enlace) ? enlace : "";
}
