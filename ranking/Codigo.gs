/**
 * Ranking familiar de Torta Runner.
 * Se pega en Extensiones → Apps Script de una hoja de Google Sheets
 * y se implementa como "Aplicación web".
 *
 * Cada partida se guarda como una fila en la pestaña "Ranking".
 * El juego muestra el mejor récord de cada nombre, sus logros
 * y el desafío de la semana (cambia los lunes).
 * El favorito de la familia de cada persona se guarda en la pestaña "Favoritos".
 */
const HOJA = 'Ranking';
const COLUMNAS = ['fecha', 'nombre', 'metros', 'frutillas', 'personaje', 'logros', 'semana', 'reto', 'valor'];
const PERSONAJES = ['sofi', 'deisy', 'leandro', 'cinthia'];
const FAVORITOS = ['sofi', 'deisy', 'leandro', 'cinthia', 'nora', 'dina'];
const HOJA_FAV = 'Favoritos';
const TOP = 20;
const SEMANA_OK = /^S\d{4}-\d{2}-\d{2}$/; // ej: S2026-10-19 (el lunes de esa semana)
const RETO_OK = /^[a-z]{3,16}$/;

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA);
    h.appendRow(COLUMNAS);
    h.setFrozenRows(1);
  } else if (String(h.getRange(1, 6).getValue()) !== 'logros') {
    // hoja creada con la versión anterior: se agregan los títulos de las columnas nuevas
    h.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS]);
  }
  return h;
}

function hojaFav_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_FAV);
  if (!h) {
    h = ss.insertSheet(HOJA_FAV);
    h.appendRow(['nombre', 'favorito', 'actualizado']);
    h.setFrozenRows(1);
  }
  return h;
}

/** Guarda (o cambia) el favorito de una persona: una fila por nombre. */
function guardarFavorito_(nombre, fav) {
  const h = hojaFav_();
  const filas = h.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).toLowerCase() === nombre.toLowerCase()) {
      h.getRange(i + 1, 2, 1, 2).setValues([[fav, new Date()]]);
      return;
    }
  }
  h.appendRow([nombre, fav, new Date()]);
}

function limpiarNombre_(s) {
  return String(s || '')
    .replace(/[<>"]/g, '')
    .replace(/^[=+\-@\s]+/, '') // evita que la hoja lo tome como fórmula
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 16);
}

function numero_(v, max) {
  const n = Math.floor(Number(v) || 0);
  return Math.max(0, Math.min(max, n));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** El juego manda una partida terminada. */
function doPost(e) {
  let d = {};
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  const nombre = limpiarNombre_(d.nombre);
  if (!nombre) return json_({ ok: false, error: 'falta el nombre' });
  const fav = FAVORITOS.indexOf(d.avatar) >= 0 ? d.avatar : '';
  if (d.tipo === 'avatar') {
    if (!fav) return json_({ ok: false, error: 'favorito no válido' });
    const l = LockService.getScriptLock();
    l.waitLock(5000);
    try { guardarFavorito_(nombre, fav); } finally { l.releaseLock(); }
    return json_({ ok: true });
  }
  const semana = SEMANA_OK.test(String(d.semana || '')) ? String(d.semana) : '';
  const reto = semana && RETO_OK.test(String(d.reto || '')) ? String(d.reto) : '';
  const fila = [
    new Date(),
    nombre,
    numero_(d.metros, 100000),
    numero_(d.frutillas, 5000),
    PERSONAJES.indexOf(d.personaje) >= 0 ? d.personaje : 'sofi',
    numero_(d.logros, 100),
    semana,
    reto,
    reto ? numero_(d.valor, 100000) : '',
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    hoja_().appendRow(fila);
    if (fav) guardarFavorito_(nombre, fav);
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

/**
 * El juego pide el ranking: mejor récord por nombre, de mayor a menor,
 * con la mayor cantidad de logros de cada uno.
 * Si manda ?semana=S2026-10-19&reto=pancitos, también devuelve el desafío de esa semana.
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const semana = SEMANA_OK.test(String(p.semana || '')) ? String(p.semana) : '';
  const reto = RETO_OK.test(String(p.reto || '')) ? String(p.reto) : '';
  const filas = hoja_().getDataRange().getValues().slice(1);
  const mejor = {};
  const logros = {};
  const sem = {};
  filas.forEach(function (f) {
    const nombre = String(f[1] || '').trim();
    if (!nombre) return;
    const k = nombre.toLowerCase();
    const metros = Number(f[2]) || 0;
    if (!mejor[k] || metros > mejor[k].metros) {
      mejor[k] = { nombre: nombre, metros: metros, frutillas: Number(f[3]) || 0, personaje: String(f[4] || 'sofi') };
    }
    logros[k] = Math.max(logros[k] || 0, Number(f[5]) || 0);
    if (semana && reto && String(f[6]) === semana && String(f[7]) === reto) {
      const valor = Number(f[8]) || 0;
      if (!sem[k] || valor > sem[k].valor) {
        sem[k] = { nombre: nombre, valor: valor, personaje: String(f[4] || 'sofi') };
      }
    }
  });
  const favs = {};
  hojaFav_().getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[0]) favs[String(f[0]).toLowerCase()] = String(f[1] || '');
  });
  Object.keys(mejor).forEach(function (k) { mejor[k].avatar = favs[k] || ''; mejor[k].logros = logros[k] || 0; });
  Object.keys(sem).forEach(function (k) { sem[k].avatar = favs[k] || ''; sem[k].logros = logros[k] || 0; });
  const top = Object.keys(mejor).map(function (k) { return mejor[k]; })
    .sort(function (a, b) { return b.metros - a.metros; })
    .slice(0, TOP);
  const res = { ok: true, top: top };
  if (semana && reto) {
    res.semana = Object.keys(sem).map(function (k) { return sem[k]; })
      .sort(function (a, b) { return b.valor - a.valor; })
      .slice(0, TOP);
  }
  return json_(res);
}
