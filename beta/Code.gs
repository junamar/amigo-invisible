/**
 * @OnlyCurrentDoc
 *
 * Almacén de notas compartidas (ideas de regalo).
 *
 * - Solo puede leer y escribir en ESTE Sheet (@OnlyCurrentDoc).
 * - Solo acepta a las personas de la pestaña "Participantes".
 * - Cada persona tiene una clave secreta dentro de su enlace: solo ese enlace
 *   puede leer las ideas y cambiar las suyas.
 * - NUNCA guarda quién regala a quién.
 */

const HOJA_IDEAS = 'Ideas de regalo';
const HOJA_PART = 'Participantes';
const MAX_TEXTO = 1500;

// ---------- utilidades ----------
function libro_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function hojaIdeas_() {
  let sh = libro_().getSheetByName(HOJA_IDEAS);
  if (!sh) {
    sh = libro_().insertSheet(HOJA_IDEAS);
    sh.appendRow(['Clave', 'Nombre', 'Ideas de regalo', 'Actualizado']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    sh.setColumnWidth(3, 420);
  }
  return sh;
}

function hojaPart_() {
  let sh = libro_().getSheetByName(HOJA_PART);
  if (!sh) {
    sh = libro_().insertSheet(HOJA_PART, 0);
    sh.appendRow(['Nombre (igual que en el sorteo)', 'Huella del enlace (automática: bórrala para reiniciar a esa persona)']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    sh.setColumnWidth(1, 260);
    sh.setColumnWidth(2, 420);
  }
  return sh;
}

// Misma conversión que la web: "Begoña" -> "begona"
function slug_(n) {
  return String(n || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

function huella_(token) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'amigo-invisible:' + token);
  return Utilities.base64EncodeWebSafe(bytes);
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

/**
 * Comprueba que la persona está en "Participantes" y que su enlace es el suyo.
 * La primera vez que alguien usa su enlace, se guarda la huella de su clave.
 * Devuelve '' si todo bien, o el código de error.
 */
function autorizar_(clave, token) {
  clave = slug_(clave);
  token = String(token || '');
  if (!clave || token.length < 8) return 'faltan_datos';
  const sh = hojaPart_();
  const filas = sh.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (slug_(filas[i][0]) !== clave) continue;
    const guardada = String(filas[i][1] || '');
    const h = huella_(token);
    if (!guardada) { sh.getRange(i + 1, 2).setValue(h); return ''; }
    return guardada === h ? '' : 'enlace_no_valido';
  }
  return 'no_participante';
}

// ---------- API ----------
// Lectura: devuelve TODAS las ideas (así nunca se sabe de quién mira cada uno).
function doGet(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const err = autorizar_(p.clave, p.tok);
    if (err) return json_({ ok: false, error: err });
    const filas = hojaIdeas_().getDataRange().getValues();
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
  } finally {
    lock.releaseLock();
  }
}

// Escritura: cada persona solo puede cambiar SUS ideas.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const err = autorizar_(d.clave, d.tok);
    if (err) return json_({ ok: false, error: err });

    const clave = slug_(d.clave);
    const nombre = seguro_(String(d.nombre || '').slice(0, 60));
    const texto = seguro_(String(d.texto || '').slice(0, MAX_TEXTO));
    const sh = hojaIdeas_();
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
    return json_({ ok: false, error: 'error' });
  } finally {
    lock.releaseLock();
  }
}

// Ejecútala una vez desde el editor (▶ Ejecutar) para crear la pestaña "Participantes".
function preparar() {
  hojaPart_();
  hojaIdeas_();
}
