/**
 * Amigo Invisible de Reyes — almacén de ideas de regalo.
 *
 * Solo guarda las ideas de regalo de cada persona (por nombre).
 * NUNCA guarda quién regala a quién: eso solo vive en los enlaces secretos.
 *
 * Instalación: pégalo en Extensiones → Apps Script de un Google Sheet vacío
 * e impleméntalo como "Aplicación web" (Ejecutar como: Yo · Acceso: Cualquier usuario).
 */

const HOJA = 'Ideas de regalo';
const MAX_TEXTO = 1500;

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(['Clave', 'Nombre', 'Ideas de regalo', 'Actualizado']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    sh.setColumnWidth(3, 420);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Evita que un texto que empiece por = + - @ se interprete como fórmula
function seguro_(txt) {
  txt = String(txt == null ? '' : txt);
  return /^[=+\-@]/.test(txt) ? "'" + txt : txt;
}

// Devuelve TODAS las ideas: así el Sheet nunca sabe de quién está mirando cada uno.
function doGet() {
  const filas = hoja_().getDataRange().getValues();
  const ideas = {};
  for (let i = 1; i < filas.length; i++) {
    const [clave, nombre, texto, fecha] = filas[i];
    if (!clave) continue;
    ideas[clave] = {
      nombre: String(nombre),
      texto: String(texto),
      fecha: fecha instanceof Date ? fecha.toISOString() : ''
    };
  }
  return json_({ ok: true, ideas: ideas });
}

// Guarda (o actualiza) las ideas de una persona.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const clave = String(d.clave || '').toLowerCase().replace(/[^a-z0-9-]/g, '').slice(0, 60);
    if (!clave) return json_({ ok: false, error: 'Falta el nombre' });
    const nombre = seguro_(String(d.nombre || '').slice(0, 60));
    const texto = seguro_(String(d.texto || '').slice(0, MAX_TEXTO));

    const sh = hoja_();
    const filas = sh.getDataRange().getValues();
    const valores = [[clave, nombre, texto, new Date()]];
    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0] === clave) {
        sh.getRange(i + 1, 1, 1, 4).setValues(valores);
        return json_({ ok: true });
      }
    }
    sh.appendRow(valores[0]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
