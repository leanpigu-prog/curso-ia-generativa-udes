// ================================================================
// Apps Script — Evidencia de uso · Curso IA Generativa UDES 2026
// ================================================================
// PASOS PARA ACTIVAR:
// 1. Crea una Google Sheet nueva llamada "Evidencia_IA_UDES_2026".
//    - Renombra la primera hoja como "Datos"
//    - Fila 1: timestamp | nombre | correo | sede | herramienta | descripcion | archivo_nombre | archivo_url
//      (el script la crea solo si falta, pero puedes dejarla puesta)
// 2. Crea una carpeta en Google Drive para guardar los archivos de evidencia
//    (ej. "Evidencias IA UDES 2026") y copia su ID (está en la URL:
//    drive.google.com/drive/folders/ESTE_ID).
// 3. Copia el ID de la Sheet (está en la URL: .../d/ESTE_ID/edit) y pégalo
//    abajo en SHEET_ID. Pega el ID de la carpeta en FOLDER_ID.
// 4. Abre Extensiones > Apps Script y pega este código completo.
// 5. Guarda (Ctrl+S).
// 6. Implementar > Nueva implementación
//    Tipo: Aplicación web | Ejecutar como: Yo | Acceso: Cualquier persona
// 7. Autoriza los permisos cuando los pida (Sheets + Drive).
// 8. Copia la URL de la implementación (empieza con https://script.google.com/macros/s/...)
//    y pégala en inicio.html como valor de EVIDENCIA_APPS_SCRIPT_URL.
// ================================================================

const SHEET_ID  = '1ddBbevAXndIvj4kQRBEKd-mh1yMM6Pb7ejK5vGTVPwg';
const FOLDER_ID = '1SvxhgcO20DKlAkhQq3eq_f65jykXVt-4';
const SHEET_NAME = 'Datos';
const HEADERS = ['timestamp', 'nombre', 'correo', 'sede', 'herramienta', 'descripcion', 'archivo_nombre', 'archivo_url'];
const MAX_BYTES = 5 * 1024 * 1024; // 5 MB, debe coincidir con EVID_MAX_BYTES en inicio.html

// POST body: {action:'guardarEvidencia', nombre, correo, sede, herramienta,
//             descripcion, archivo_nombre, archivo_tipo, archivo_base64}
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action !== 'guardarEvidencia') return resp({ ok: false, error: 'acción desconocida' });

    const requeridos = ['nombre', 'correo', 'sede', 'herramienta', 'descripcion', 'archivo_nombre', 'archivo_base64'];
    for (const campo of requeridos) {
      if (!p[campo]) return resp({ ok: false, error: 'falta el campo ' + campo });
    }

    const bytes = Utilities.base64Decode(p.archivo_base64);
    if (bytes.length > MAX_BYTES) return resp({ ok: false, error: 'el archivo supera el límite permitido' });

    const blob = Utilities.newBlob(bytes, p.archivo_tipo || 'application/octet-stream', p.archivo_nombre);
    const carpeta = DriveApp.getFolderById(FOLDER_ID);
    const archivo = carpeta.createFile(blob);
    archivo.setName(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + p.archivo_nombre);

    const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (ws.getLastRow() === 0) {
      ws.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }

    ws.appendRow([
      new Date().toISOString(),
      p.nombre,
      p.correo,
      p.sede,
      p.herramienta,
      p.descripcion,
      archivo.getName(),
      archivo.getUrl()
    ]);

    return resp({ ok: true });
  } catch (err) {
    return resp({ ok: false, error: err.message });
  }
}

function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
