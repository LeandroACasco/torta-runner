/**
 * Ranking familiar de Torta Runner.
 * Se pega en Extensiones → Apps Script de una hoja de Google Sheets
 * y se implementa como "Aplicación web".
 *
 * Cada partida se guarda como una fila en la pestaña "Ranking".
 * El juego muestra el mejor récord de cada nombre, sus logros
 * y el desafío de la semana (cambia los lunes).
 * El favorito de la familia de cada persona se guarda en la pestaña "Favoritos".
 * La caja fuerte (progreso de cada persona) se guarda en la pestaña "Cuentas".
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

/** El juego manda una partida terminada (o una acción de la caja fuerte). */
function doPost(e) {
  let d = {};
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  if (d.tipo === 'cuenta') return json_(cuenta_(d));
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
  const anterior = SEMANA_OK.test(String(p.anterior || '')) ? String(p.anterior) : '';
  const filas = hoja_().getDataRange().getValues().slice(1);
  const mejor = {};
  const logros = {};
  const sem = {};
  const semanas = {}; // mejor valor de cada persona en cada semana (para saber quién ganó)
  filas.forEach(function (f) {
    const nombre = String(f[1] || '').trim();
    if (!nombre) return;
    const k = nombre.toLowerCase();
    const metros = Number(f[2]) || 0;
    if (!mejor[k] || metros > mejor[k].metros) {
      mejor[k] = { nombre: nombre, metros: metros, frutillas: Number(f[3]) || 0, personaje: String(f[4] || 'sofi') };
    }
    logros[k] = Math.max(logros[k] || 0, Number(f[5]) || 0);
    const fs = String(f[6] || '');
    if (SEMANA_OK.test(fs) && f[7] !== '' && f[8] !== '') {
      const v = Number(f[8]) || 0;
      const g = fs + '|' + String(f[7]); // cada semana tiene un solo desafío; se separan por las dudas
      semanas[g] = semanas[g] || {};
      if (!semanas[g][k] || v > semanas[g][k].valor) semanas[g][k] = { nombre: nombre, valor: v, reto: String(f[7]) };
    }
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
  // semanas ganadas: solo las que ya terminaron (anteriores a la semana que pide el juego)
  if (semana) {
    const campeones = {};
    let ganador = null, filasGanador = 0;
    Object.keys(semanas).forEach(function (gk) {
      const s = gk.split('|')[0];
      if (s >= semana) return;
      let g = null;
      const quienes = Object.keys(semanas[gk]);
      quienes.forEach(function (k) { if (!g || semanas[gk][k].valor > g.valor) g = Object.assign({ k: k }, semanas[gk][k]); });
      if (!g || g.valor <= 0) return;
      campeones[g.k] = (campeones[g.k] || 0) + 1;
      if (s === anterior && quienes.length > filasGanador) {
        filasGanador = quienes.length;
        ganador = { nombre: g.nombre, valor: g.valor, reto: g.reto, avatar: favs[g.k] || '' };
      }
    });
    res.campeones = campeones;
    res.anterior = ganador;
  }
  if (semana && reto) {
    res.semana = Object.keys(sem).map(function (k) { return sem[k]; })
      .sort(function (a, b) { return b.valor - a.valor; })
      .slice(0, TOP);
  }
  return json_(res);
}

/* =====================================================================
 * Caja fuerte: el progreso de cada persona guardado en la nube.
 * Pestaña "Cuentas", una fila por persona. El PIN y las respuestas
 * de las preguntas secretas se guardan cifrados (hash con sal), nunca
 * tal cual. La "llave" es lo que cada dispositivo guarda para
 * sincronizar sin volver a pedir el PIN; cambia si se cambia el PIN.
 * ===================================================================== */
const HOJA_CUENTAS = 'Cuentas';
const COL_CUENTAS = ['nombre', 'sal', 'pin', 'llave', 'preguntas', 'respuestas', 'progreso',
  'ultimoSync', 'intentos', 'bloqueoHasta', 'creado', 'actualizado'];
const PREGUNTAS = ['quiero', 'risa', 'jugar', 'abrazos', 'cocina', 'dormilon', 'secreto', 'baila', 'viaje', 'travieso'];
const MAX_INTENTOS = 5;   // después de 5 errores seguidos…
const BLOQUEO_MIN = 15;   // …la caja fuerte espera 15 minutos

// Cómo se combina el progreso de dos dispositivos
const PROG_SUMA = ['total', 'st.runs', 'st.pans', 'st.kamps', 'st.tortas', 'st.escudos', 'st.escapes', 'st.scared', 'st.meters'];
const PROG_MAX = ['best', 'st.bestFru', 'st.bestPans', 'st.bestEscapes', 'st.cleanBest'];
const PROG_LISTA = ['seen', 'st.powers', 'st.chars'];

function hojaCuentas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA_CUENTAS);
  if (!h) {
    h = ss.insertSheet(HOJA_CUENTAS);
    h.getRange(1, 1, 1000, COL_CUENTAS.length).setNumberFormat('@'); // todo como texto
    h.appendRow(COL_CUENTAS);
    h.setFrozenRows(1);
  }
  return h;
}

function hash_(s) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
}
function azar_(p) { return p + Utilities.getUuid().replace(/-/g, ''); }
function pinOk_(p) { return /^\d{4}$/.test(String(p || '')); }

function buscarCuenta_(h, nombre) {
  const filas = h.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).toLowerCase() === nombre.toLowerCase()) {
      const c = {};
      COL_CUENTAS.forEach(function (k, j) { c[k] = filas[i][j]; });
      c.fila = i + 1;
      return c;
    }
  }
  return null;
}

function guardarCuenta_(h, c) {
  c.actualizado = new Date().toISOString();
  const fila = COL_CUENTAS.map(function (k) { return c[k] === undefined || c[k] === null ? '' : String(c[k]); });
  if (c.fila) h.getRange(c.fila, 1, 1, COL_CUENTAS.length).setValues([fila]);
  else h.appendRow(fila);
}

function leerJson_(s, def) { try { return JSON.parse(s) || def; } catch (e) { return def; } }

// ---- progreso: limpiar y combinar ----
function tomar_(o, p) { return p.split('.').reduce(function (a, k) { return a && a[k]; }, o); }
function poner_(o, p, v) {
  const ks = p.split('.'); let a = o;
  for (let i = 0; i < ks.length - 1; i++) a = a[ks[i]] = a[ks[i]] || {};
  a[ks[ks.length - 1]] = v;
}
function n_(v) { return Math.max(0, Math.min(10000000, Math.floor(Number(v) || 0))); }
function palabras_(a) {
  return (Array.isArray(a) ? a : []).map(String).filter(function (s) { return /^[a-z0-9_]{1,24}$/.test(s); }).slice(0, 40);
}

/** Deja solo los datos de progreso que el juego entiende. */
function limpiarProgreso_(p) {
  p = p || {};
  const o = { st: {} };
  PROG_SUMA.concat(PROG_MAX).forEach(function (k) { poner_(o, k, n_(tomar_(p, k))); });
  PROG_LISTA.forEach(function (k) { poner_(o, k, palabras_(tomar_(p, k))); });
  o.st.fiesta = !!(p.st && p.st.fiesta);
  o.ach = {}; o.tips = {};
  Object.keys(p.ach || {}).slice(0, 60).forEach(function (k) { if (/^[a-z0-9_]{1,24}$/.test(k)) o.ach[k] = n_(p.ach[k]) || 1; });
  Object.keys(p.tips || {}).slice(0, 40).forEach(function (k) { if (/^[a-z0-9_]{1,24}$/.test(k)) o.tips[k] = 1; });
  return o;
}

/**
 * Combina lo que hay en la nube con lo que manda un dispositivo.
 * base = lo último que ese dispositivo recibió de la nube, así se suma
 * solo lo nuevo (las frutillas de dos dispositivos se suman, no se pisan).
 */
function combinar_(nube, local, base) {
  const o = { st: {} };
  PROG_SUMA.forEach(function (k) { poner_(o, k, n_(tomar_(nube, k)) + Math.max(0, n_(tomar_(local, k)) - n_(tomar_(base, k)))); });
  PROG_MAX.forEach(function (k) { poner_(o, k, Math.max(n_(tomar_(nube, k)), n_(tomar_(local, k)))); });
  PROG_LISTA.forEach(function (k) {
    const u = (tomar_(nube, k) || []).slice();
    (tomar_(local, k) || []).forEach(function (x) { if (u.indexOf(x) < 0) u.push(x); });
    poner_(o, k, u);
  });
  o.st.fiesta = !!(nube.st && nube.st.fiesta) || !!(local.st && local.st.fiesta);
  o.ach = {};
  [nube.ach || {}, local.ach || {}].forEach(function (src) {
    Object.keys(src).forEach(function (k) { const t = n_(src[k]) || 1; o.ach[k] = o.ach[k] ? Math.min(o.ach[k], t) : t; });
  });
  o.tips = {};
  [nube.tips || {}, local.tips || {}].forEach(function (src) { Object.keys(src).forEach(function (k) { o.tips[k] = 1; }); });
  return o;
}

// ---- intentos y bloqueo ----
function bloqueada_(c) {
  const hasta = Number(c.bloqueoHasta) || 0;
  return hasta > Date.now() ? Math.ceil((hasta - Date.now()) / 60000) : 0;
}
function fallo_(h, c, error) {
  let n = (Number(c.intentos) || 0) + 1;
  if (n >= MAX_INTENTOS) { c.bloqueoHasta = Date.now() + BLOQUEO_MIN * 60000; n = 0; }
  c.intentos = n;
  guardarCuenta_(h, c);
  const min = bloqueada_(c);
  return min ? { ok: false, error: 'bloqueo', minutos: min } : { ok: false, error: error, quedan: MAX_INTENTOS - n };
}
function abierta_(h, c) {
  c.intentos = 0; c.bloqueoHasta = '';
  guardarCuenta_(h, c);
  return { ok: true, nombre: c.nombre, llave: c.llave, progreso: leerJson_(c.progreso, {}), preguntas: leerJson_(c.preguntas, []) };
}

function respuestasOk_(lista) {
  if (!Array.isArray(lista) || lista.length < 1 || lista.length > 3) return null;
  const vistas = {};
  for (let i = 0; i < lista.length; i++) {
    const q = String(lista[i] && lista[i].q || ''), a = String(lista[i] && lista[i].a || '');
    if (PREGUNTAS.indexOf(q) < 0 || FAVORITOS.indexOf(a) < 0 || vistas[q]) return null;
    vistas[q] = a;
  }
  return vistas;
}

/** Todas las acciones de la caja fuerte. */
function cuenta_(d) {
  const nombre = limpiarNombre_(d.nombre);
  if (!nombre) return { ok: false, error: 'nombre' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'ocupado' };
  try {
    const h = hojaCuentas_();
    const c = buscarCuenta_(h, nombre);
    const accion = String(d.accion || '');

    if (accion === 'preguntas') { // ¿existe? y qué preguntas eligió (para recuperar el PIN)
      return c ? { ok: true, existe: true, nombre: c.nombre, preguntas: leerJson_(c.preguntas, []) } : { ok: true, existe: false };
    }

    if (accion === 'crear') {
      if (c) return { ok: false, error: 'existe' };
      if (!pinOk_(d.pin)) return { ok: false, error: 'pin' };
      const resp = respuestasOk_(d.preguntas);
      if (!resp) return { ok: false, error: 'preguntas' };
      const sal = azar_('s');
      const qs = Object.keys(resp);
      const hashes = {};
      qs.forEach(function (q) { hashes[q] = hash_(sal + '|' + q + '|' + resp[q]); });
      const nueva = {
        nombre: nombre, sal: sal, pin: hash_(sal + '|pin|' + d.pin), llave: azar_('k'),
        preguntas: JSON.stringify(qs), respuestas: JSON.stringify(hashes),
        progreso: JSON.stringify(limpiarProgreso_(d.progreso)), ultimoSync: '', intentos: 0, bloqueoHasta: '',
        creado: new Date().toISOString(),
      };
      guardarCuenta_(h, nueva);
      return { ok: true, nombre: nombre, llave: nueva.llave, progreso: leerJson_(nueva.progreso, {}), preguntas: qs };
    }

    if (!c) return { ok: false, error: 'noexiste' };

    if (accion === 'sync') {
      if (!d.llave || String(d.llave) !== String(c.llave)) return { ok: false, error: 'llave' };
      const id = String(d.syncId || '').slice(0, 40);
      if (!id || id !== String(c.ultimoSync)) { // un reintento del mismo envío no suma dos veces
        const nube = limpiarProgreso_(leerJson_(c.progreso, {}));
        c.progreso = JSON.stringify(combinar_(nube, limpiarProgreso_(d.local), limpiarProgreso_(d.base)));
        c.ultimoSync = id;
        guardarCuenta_(h, c);
      }
      return { ok: true, progreso: leerJson_(c.progreso, {}) };
    }

    const min = bloqueada_(c);
    if (min) return { ok: false, error: 'bloqueo', minutos: min };

    if (accion === 'entrar') {
      if (!pinOk_(d.pin) || hash_(c.sal + '|pin|' + d.pin) !== String(c.pin)) return fallo_(h, c, 'pin');
      return abierta_(h, c);
    }

    if (accion === 'recuperar') { // respondió las preguntas secretas: elige un PIN nuevo
      if (!pinOk_(d.pinNuevo)) return { ok: false, error: 'pin' };
      const qs = leerJson_(c.preguntas, []), guardadas = leerJson_(c.respuestas, {});
      const resp = {};
      (Array.isArray(d.respuestas) ? d.respuestas : []).forEach(function (r) { if (r) resp[String(r.q)] = String(r.a); });
      const bien = qs.length > 0 && qs.every(function (q) { return hash_(c.sal + '|' + q + '|' + resp[q]) === guardadas[q]; });
      if (!bien) return fallo_(h, c, 'respuestas');
      c.pin = hash_(c.sal + '|pin|' + d.pinNuevo);
      c.llave = azar_('k');
      return abierta_(h, c);
    }

    if (accion === 'cambiarPin') {
      if (!pinOk_(d.pin) || hash_(c.sal + '|pin|' + d.pin) !== String(c.pin)) return fallo_(h, c, 'pin');
      if (!pinOk_(d.pinNuevo)) return { ok: false, error: 'pin' };
      c.pin = hash_(c.sal + '|pin|' + d.pinNuevo);
      c.llave = azar_('k');
      return abierta_(h, c);
    }

    return { ok: false, error: 'accion' };
  } finally {
    lock.releaseLock();
  }
}
