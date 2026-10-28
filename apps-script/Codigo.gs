/**
 * ALTUS GEAR — Caja Altus · Motor de la hoja
 * v1.0 · 2026-10-27
 *
 * Este código vive dentro de la hoja Altus_Caja_2026 (Extensiones → Apps Script).
 *   1. configurar()     — construye hojas, resumen, carpetas, clave y disparador. Se corre UNA vez.
 *   2. doPost()         — es la API que usa la app del teléfono. Toda llamada lleva la clave.
 *   3. procesarBuzon()  — aplica los archivos que Claude deja en buzon_claude/ (clasificar gastos,
 *                          completar movimientos leídos de fotos). Corre cada hora y en cada sincronización.
 *
 * REGLAS
 *   - MOVIMIENTOS no se edita a mano: escriben la app y el buzón. Lo eliminado se marca, no se borra.
 *   - CONCEPTOS.clase y METAS sí se pueden editar a mano. Un cambio de clase se propaga solo.
 */

var VERSION = '1.0';
var TZ = 'America/Mexico_City';

var MOV = ['id','tipo','fecha','mes','quien','linea','subcategoria','producto','cantidad','precio_unitario',
  'costo_unitario','costo_total','monto','canal','metodo','origen','concepto','concepto_id','clase','cargo',
  'unidad','piezas','proveedor','socio','direccion','nota','foto_url','fuente','por_revisar','por_leer',
  'estatus','creado_en','actualizado_en','dispositivo'];
var CONC = ['concepto_id','nombre','clase','creado_en','clasificado_en','clasificado_por'];
var BIT = ['fecha','origen','accion','detalle'];

/** Campos que la app o el buzón pueden escribir. Los demás (mes, concepto_id, clase, costo_total, estatus) los calcula el motor. */
var CAMPOS = ['tipo','fecha','quien','linea','subcategoria','producto','cantidad','precio_unitario','costo_unitario',
  'monto','canal','metodo','origen','concepto','cargo','unidad','piezas','proveedor','socio','direccion','nota',
  'foto_url','fuente','por_revisar','por_leer','creado_en','actualizado_en','dispositivo'];
var NUMEROS = ['cantidad','precio_unitario','costo_unitario','monto','piezas'];
var TEXTO_COLS = ['id','tipo','fecha','mes','quien','linea','subcategoria','producto','canal','metodo','origen',
  'concepto','concepto_id','clase','cargo','unidad','proveedor','socio','direccion','nota','foto_url','fuente',
  'estatus','creado_en','actualizado_en','dispositivo'];
var TIPOS = ['venta','gasto','compra','socios'];
var CLASES = ['fijo','semifijo','variable'];

var CATALOGO_DEF = {
  lineas: {
    Apparel: ['Playeras','Sudaderas','Shorts','Accesorios'],
    Calzado: ['Senderismo','Correr','Entrenamiento','Casual','Otro'],
    Management: ['Retainers','Producción','Activaciones','Comisiones y licencias']
  },
  canales: ['Mercado','Evento','WhatsApp','Tienda en línea','Directo / B2B'],
  origenes: ['TikTok','Instagram','Recomendación','Atleta','Cliente recurrente','Pasaba por ahí'],
  metodos: ['Efectivo','Transferencia','Tarjeta','Otro'],
  personas: ['Fernando','Alberto'],
  split: { Apparel: 0.2333, Calzado: 0.4667, Management: 0.30 }
};
/** Business Plan 2026 §7 (vía negocio.md). Calzado unificado según D010. */
var METAS_2026 = [['2026-01',90000],['2026-02',90000],['2026-03',105000],['2026-04',105000],['2026-05',120000],
  ['2026-06',120000],['2026-07',120000],['2026-08',120000],['2026-09',135000],['2026-10',150000],
  ['2026-11',165000],['2026-12',180000]];

// ─────────────────────────────────────────────────────────────
// MENÚ
// ─────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Altus Caja')
    .addItem('Configurar (primera vez)', 'configurar')
    .addItem('Procesar buzón de Claude ahora', 'procesarBuzonUI')
    .addItem('Ver clave de la app', 'verClave')
    .addItem('Reconstruir resumen', 'construirResumen')
    .addToUi();
}

/** Si alguien cambia la clase de un concepto a mano, se propaga a sus movimientos. */
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== 'CONCEPTOS' || e.range.getRow() < 2) return;
    var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (e.range.getColumn() !== H.indexOf('clase') + 1) return;
    var id = sh.getRange(e.range.getRow(), H.indexOf('concepto_id') + 1).getValue();
    var clase = String(e.range.getValue() || '').toLowerCase();
    if (clase && CLASES.indexOf(clase) < 0) return;
    sh.getRange(e.range.getRow(), H.indexOf('clasificado_en') + 1).setValue(iso_());
    sh.getRange(e.range.getRow(), H.indexOf('clasificado_por') + 1).setValue('hoja');
    propagarClase_(id, clase);
  } catch (err) {}
}

// ─────────────────────────────────────────────────────────────
// CONFIGURAR (una vez)
// ─────────────────────────────────────────────────────────────

function configurar() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  var mov = hoja_(ss, 'MOVIMIENTOS', MOV);
  TEXTO_COLS.forEach(function (c) { mov.getRange(1, MOV.indexOf(c) + 1, mov.getMaxRows(), 1).setNumberFormat('@'); });
  NUMEROS.concat(['costo_total']).forEach(function (c) {
    mov.getRange(2, MOV.indexOf(c) + 1, mov.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');
  });

  var conc = hoja_(ss, 'CONCEPTOS', CONC);
  conc.getRange(1, 1, conc.getMaxRows(), CONC.length).setNumberFormat('@');
  var regla = SpreadsheetApp.newDataValidation().requireValueInList(CLASES, true).setAllowInvalid(false).build();
  conc.getRange(2, CONC.indexOf('clase') + 1, conc.getMaxRows() - 1, 1).setDataValidation(regla);

  var cat = hoja_(ss, 'CATALOGO', ['clave', 'valor_json']);
  if (cat.getLastRow() < 2) {
    Object.keys(CATALOGO_DEF).forEach(function (k) { cat.appendRow([k, JSON.stringify(CATALOGO_DEF[k])]); });
  }

  var metas = hoja_(ss, 'METAS', ['mes', 'meta_total']);
  metas.getRange(1, 1, metas.getMaxRows(), 1).setNumberFormat('@');
  if (metas.getLastRow() < 2) metas.getRange(2, 1, METAS_2026.length, 2).setValues(METAS_2026);
  metas.getRange(2, 2, metas.getMaxRows() - 1, 1).setNumberFormat('$#,##0');

  hoja_(ss, 'BITACORA', BIT);
  construirResumen();
  carpetas_();

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CLAVE')) props.setProperty('CLAVE', nuevaClave_());

  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'procesarBuzon'; });
  if (!hay) ScriptApp.newTrigger('procesarBuzon').timeBased().everyHours(1).create();

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  log_('motor', 'configurar', 'v' + VERSION);
  SpreadsheetApp.getUi().alert(
    'Caja Altus lista',
    'Clave de la app: ' + props.getProperty('CLAVE') + '\n\n' +
    'Guárdala: se escribe una sola vez en cada teléfono.\n\n' +
    'Siguiente paso: Implementar → Nueva implementación → Aplicación web.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function verClave() {
  var c = PropertiesService.getScriptProperties().getProperty('CLAVE') || '(sin clave: corre Configurar)';
  SpreadsheetApp.getUi().alert('Clave de la app', c, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Hoja RESUMEN: P&L preliminar por mes con fórmulas sobre MOVIMIENTOS. */
function construirResumen() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('RESUMEN') || ss.insertSheet('RESUMEN');
  sh.clear();
  var cols = ['mes','meta','ventas','apparel','calzado','management','avance_meta','costo_vendido','ventas_sin_costo',
    'margen_bruto','gastos_fijos','gastos_semifijos','gastos_variables','gastos_por_clasificar','resultado_operativo',
    'compras_inventario','aportaciones','retiros'];
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  var L = function (name) { return 'MOVIMIENTOS!$' + col_(MOV.indexOf(name) + 1) + ':$' + col_(MOV.indexOf(name) + 1); };
  var base = function (r) { return L('mes') + ',$A' + r + ',' + L('estatus') + ',"<>eliminado"'; };
  var suma = function (r, extra) { return '=SUMIFS(' + L('monto') + ',' + base(r) + (extra ? ',' + extra : '') + ')'; };
  var rows = [];
  for (var i = 0; i < METAS_2026.length; i++) {
    var r = i + 2;
    rows.push([
      METAS_2026[i][0],
      '=IFERROR(VLOOKUP($A' + r + ',METAS!$A:$B,2,FALSE),0)',
      suma(r, L('tipo') + ',"venta"'),
      suma(r, L('tipo') + ',"venta",' + L('linea') + ',"Apparel"'),
      suma(r, L('tipo') + ',"venta",' + L('linea') + ',"Calzado"'),
      suma(r, L('tipo') + ',"venta",' + L('linea') + ',"Management"'),
      '=IF($B' + r + '>0,$C' + r + '/$B' + r + ',"")',
      '=SUMIFS(' + L('costo_total') + ',' + base(r) + ',' + L('tipo') + ',"venta")',
      suma(r, L('tipo') + ',"venta",' + L('costo_total') + ',""'),
      '=$C' + r + '-$I' + r + '-$H' + r,
      suma(r, L('tipo') + ',"gasto",' + L('clase') + ',"fijo"'),
      suma(r, L('tipo') + ',"gasto",' + L('clase') + ',"semifijo"'),
      suma(r, L('tipo') + ',"gasto",' + L('clase') + ',"variable"'),
      suma(r, L('tipo') + ',"gasto",' + L('clase') + ',""'),
      '=$C' + r + '-$H' + r + '-$K' + r + '-$L' + r + '-$M' + r + '-$N' + r,
      suma(r, L('tipo') + ',"compra"'),
      suma(r, L('tipo') + ',"socios",' + L('direccion') + ',"aportacion"'),
      suma(r, L('tipo') + ',"socios",' + L('direccion') + ',"retiro"')
    ]);
  }
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
  sh.getRange(2, 2, rows.length, cols.length - 1).setNumberFormat('$#,##0');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(rows.length + 3, 1).setValue(
    'Notas: margen_bruto es sobre ventas con costo conocido. resultado_operativo queda sobrestimado mientras haya ventas_sin_costo. ' +
    'Compras de inventario y movimientos de socios no entran al resultado.');
}

// ─────────────────────────────────────────────────────────────
// API (la usa la app)
// ─────────────────────────────────────────────────────────────

function doGet() {
  return ContentService.createTextOutput('Caja Altus · motor v' + VERSION + ' activo.');
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    var clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
    if (!clave || String(p.clave || '').trim().toUpperCase() !== clave) return out_({ ok: false, error: 'clave' });
    switch (p.accion) {
      case 'ping': return out_({ ok: true, version: VERSION });
      case 'datos': return out_(Object.assign({ ok: true, version: VERSION }, datos_()));
      case 'guardar': return out_({ ok: true, guardados: guardar_(p.ops || [], 'app') });
      case 'foto': return out_({ ok: true, url: foto_(p) });
      case 'clasificar': clasificar_(p.concepto_id, p.clase, p.quien || 'app'); return out_({ ok: true });
    }
    return out_({ ok: false, error: 'accion' });
  } catch (err) {
    log_('api', 'error', String(err && err.message || err));
    return out_({ ok: false, error: 'servidor', mensaje: String(err && err.message || err) });
  }
}

function datos_() {
  try { procesarBuzon_(); } catch (err) { log_('buzon', 'error', String(err)); }
  var ss = SpreadsheetApp.getActive();
  var movimientos = filas_(ss.getSheetByName('MOVIMIENTOS')).filter(function (m) { return m.estatus !== 'eliminado' && m.id; });
  var conceptos = filas_(ss.getSheetByName('CONCEPTOS')).filter(function (c) { return c.concepto_id; });
  var catalogo = {};
  filas_(ss.getSheetByName('CATALOGO')).forEach(function (r) {
    try { catalogo[r.clave] = JSON.parse(r.valor_json); } catch (err) {}
  });
  var metas = {};
  filas_(ss.getSheetByName('METAS')).forEach(function (r) { if (r.mes) metas[String(r.mes)] = Number(r.meta_total) || 0; });
  catalogo.metas = metas;
  return { movimientos: movimientos, conceptos: conceptos, catalogo: catalogo };
}

/** ops: [{op:'set'|'del', id, data}] */
function guardar_(ops, origen) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('MOVIMIENTOS');
    var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, H.indexOf('id') + 1, sh.getLastRow() - 1, 1).getValues() : [];
    var fila = {};
    ids.forEach(function (v, i) { if (v[0]) fila[String(v[0])] = i + 2; });
    var conc = conceptosMapa_();
    var hechos = [];
    ops.forEach(function (op) {
      var id = String(op.id || '');
      if (!/^[A-Za-z0-9_-]{4,64}$/.test(id)) throw new Error('id inválido');
      if (op.op === 'del') {
        if (fila[id]) {
          sh.getRange(fila[id], H.indexOf('estatus') + 1).setValue('eliminado');
          sh.getRange(fila[id], H.indexOf('actualizado_en') + 1).setValue(iso_());
          log_(origen, 'eliminar', id);
        }
        hechos.push(id);
        return;
      }
      var d = limpiar_(op.data || {});
      if (TIPOS.indexOf(d.tipo) < 0) throw new Error('tipo inválido en ' + id);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.fecha || ''))) throw new Error('fecha inválida en ' + id);
      d.id = id;
      derivar_(d, conc);
      if (fila[id]) {
        var prev = filaObjeto_(sh, H, fila[id]);
        if (prev.creado_en && !d.creado_en) d.creado_en = prev.creado_en;
        if (prev.foto_url && !d.foto_url) d.foto_url = prev.foto_url;
        sh.getRange(fila[id], 1, 1, H.length).setValues([H.map(function (h) { return valor_(d[h]); })]);
      } else {
        sh.appendRow(H.map(function (h) { return valor_(d[h]); }));
        fila[id] = sh.getLastRow();
      }
      hechos.push(id);
    });
    return hechos;
  } finally {
    lock.releaseLock();
  }
}

function foto_(p) {
  var id = String(p.id || '');
  if (!/^[A-Za-z0-9_-]{4,64}$/.test(id)) throw new Error('id inválido');
  var bytes = Utilities.base64Decode(String(p.base64 || ''));
  if (!bytes.length) throw new Error('foto vacía');
  if (bytes.length > 8 * 1024 * 1024) throw new Error('foto demasiado grande');
  var fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(p.fecha || '')) ? p.fecha : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var f = carpetas_().fotos.createFile(Utilities.newBlob(bytes, 'image/jpeg', fecha + '_' + id + '.jpg'));
  return f.getUrl();
}

function clasificar_(conceptoId, clase, por) {
  conceptoId = String(conceptoId || '');
  clase = String(clase || '').toLowerCase();
  if (!conceptoId || (clase && CLASES.indexOf(clase) < 0)) throw new Error('clasificación inválida');
  var sh = SpreadsheetApp.getActive().getSheetByName('CONCEPTOS');
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, H.length).getValues() : [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][H.indexOf('concepto_id')]) === conceptoId) {
      sh.getRange(i + 2, H.indexOf('clase') + 1).setValue(clase);
      sh.getRange(i + 2, H.indexOf('clasificado_en') + 1).setValue(iso_());
      sh.getRange(i + 2, H.indexOf('clasificado_por') + 1).setValue(por);
      propagarClase_(conceptoId, clase);
      log_(por, 'clasificar', conceptoId + ' → ' + clase);
      return;
    }
  }
  throw new Error('concepto no encontrado: ' + conceptoId);
}

// ─────────────────────────────────────────────────────────────
// BUZÓN DE CLAUDE
// Archivos JSON (o texto con JSON) en buzon_claude/, con la forma:
//   {"acciones":[
//     {"tipo":"clasificar","concepto_id":"renta-del-puesto","clase":"fijo"},
//     {"tipo":"completar","id":"m...","campos":{"monto":450,"concepto":"Envío","por_leer":false}}
//   ]}
// Cada archivo aplicado se mueve a buzon_claude/procesados/.
// ─────────────────────────────────────────────────────────────

function procesarBuzon() { procesarBuzon_(); }

function procesarBuzonUI() {
  var n = procesarBuzon_();
  SpreadsheetApp.getUi().alert('Buzón de Claude', n + ' archivo(s) aplicados.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function procesarBuzon_() {
  var c = carpetas_();
  var it = c.buzon.getFiles();
  var n = 0;
  while (it.hasNext()) {
    var f = it.next();
    var texto;
    try {
      texto = f.getMimeType() === MimeType.GOOGLE_DOCS ? DocumentApp.openById(f.getId()).getBody().getText() : f.getBlob().getDataAsString('UTF-8');
      var ini = texto.indexOf('{'), fin = texto.lastIndexOf('}');
      var obj = JSON.parse(texto.slice(ini, fin + 1));
      var acciones = obj.acciones || [];
      var conc = conceptosMapa_();
      var ops = [];
      acciones.forEach(function (a) {
        if (a.tipo === 'clasificar') {
          var cid = a.concepto_id || slug_(a.concepto || '');
          if (!conc[cid]) altaConcepto_(cid, a.concepto || cid);
          clasificar_(cid, a.clase, 'claude');
        } else if (a.tipo === 'completar' && a.id) {
          var actual = movimientoPorId_(a.id);
          if (!actual) return;
          var campos = limpiar_(a.campos || {});
          Object.keys(campos).forEach(function (k) { actual[k] = campos[k]; });
          actual.actualizado_en = iso_();
          actual.fuente = actual.fuente || 'foto';
          ops.push({ op: 'set', id: a.id, data: actual });
        }
      });
      if (ops.length) guardar_(ops, 'claude');
      f.moveTo(c.procesados);
      log_('claude', 'buzon', f.getName() + ' · ' + acciones.length + ' acciones');
      n++;
    } catch (err) {
      log_('claude', 'buzon_error', f.getName() + ' · ' + String(err && err.message || err));
      f.setName('ERROR_' + f.getName().replace(/^ERROR_/, ''));
      f.moveTo(c.procesados);
    }
  }
  return n;
}

// ─────────────────────────────────────────────────────────────
// AUXILIARES
// ─────────────────────────────────────────────────────────────

function derivar_(d, conc) {
  d.mes = String(d.fecha).slice(0, 7);
  d.estatus = 'activo';
  if (d.tipo === 'venta') {
    var q = num_(d.cantidad), p = num_(d.precio_unitario), cu = num_(d.costo_unitario);
    if (num_(d.monto) == null && q != null && p != null) d.monto = Math.round(q * p * 100) / 100;
    d.costo_total = (cu != null && q != null) ? Math.round(q * cu * 100) / 100 : '';
    d.unidad = d.linea === 'Management' ? 'Management' : 'Gear';
  } else {
    d.costo_total = '';
  }
  if (d.tipo === 'gasto' && d.concepto) {
    d.concepto_id = slug_(d.concepto);
    if (!conc[d.concepto_id]) { altaConcepto_(d.concepto_id, d.concepto); conc[d.concepto_id] = { clase: '' }; }
    d.clase = conc[d.concepto_id].clase || '';
    if (!d.cargo) d.cargo = 'General';
    d.unidad = d.cargo === 'Management' ? 'Management' : (d.cargo === 'General' ? 'General' : 'Gear');
  } else {
    d.concepto_id = ''; d.clase = '';
  }
  if (d.tipo === 'compra') d.unidad = 'Gear';
  if (d.tipo === 'socios' && !d.socio) d.socio = d.quien || '';
}

function limpiar_(src) {
  var d = {};
  CAMPOS.forEach(function (k) {
    if (!(k in src)) return;
    var v = src[k];
    if (NUMEROS.indexOf(k) >= 0) { d[k] = num_(v); return; }
    if (k === 'por_revisar' || k === 'por_leer') { d[k] = v === true || v === 'true' || v === 'TRUE'; return; }
    if (v === null || v === undefined) { d[k] = ''; return; }
    var s = String(v).trim().slice(0, 500);
    if (/^[=+\-@]/.test(s)) s = "'" + s;
    d[k] = s;
  });
  return d;
}

function valor_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v;
  return v;
}

function propagarClase_(conceptoId, clase) {
  var sh = SpreadsheetApp.getActive().getSheetByName('MOVIMIENTOS');
  if (sh.getLastRow() < 2) return;
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var ci = H.indexOf('concepto_id'), cl = H.indexOf('clase');
  var rng = sh.getRange(2, 1, sh.getLastRow() - 1, H.length);
  var vals = rng.getValues();
  var cambio = false;
  vals.forEach(function (r) { if (String(r[ci]) === String(conceptoId)) { r[cl] = clase; cambio = true; } });
  if (cambio) sh.getRange(2, cl + 1, vals.length, 1).setValues(vals.map(function (r) { return [r[cl]]; }));
}

function conceptosMapa_() {
  var m = {};
  filas_(SpreadsheetApp.getActive().getSheetByName('CONCEPTOS')).forEach(function (c) {
    if (c.concepto_id) m[c.concepto_id] = { nombre: c.nombre, clase: String(c.clase || '') };
  });
  return m;
}

function altaConcepto_(id, nombre) {
  var sh = SpreadsheetApp.getActive().getSheetByName('CONCEPTOS');
  sh.appendRow(CONC.map(function (h) { return h === 'concepto_id' ? id : h === 'nombre' ? nombre : h === 'creado_en' ? iso_() : ''; }));
}

function movimientoPorId_(id) {
  var sh = SpreadsheetApp.getActive().getSheetByName('MOVIMIENTOS');
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (sh.getLastRow() < 2) return null;
  var ids = sh.getRange(2, H.indexOf('id') + 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return filaObjeto_(sh, H, i + 2);
  return null;
}

function filaObjeto_(sh, H, row) {
  var v = sh.getRange(row, 1, 1, H.length).getValues()[0];
  var o = {};
  H.forEach(function (h, i) { o[h] = normal_(h, v[i]); });
  return o;
}

function filas_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var H = vals.shift();
  return vals.map(function (r) {
    var o = {};
    H.forEach(function (h, i) { if (h) o[h] = normal_(h, r[i]); });
    return o;
  });
}

function normal_(h, v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, h === 'mes' ? 'yyyy-MM' : (h === 'fecha' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss"));
  if (v === '' && NUMEROS.indexOf(h) >= 0) return null;
  if (typeof v === 'string' && v.charAt(0) === "'") return v.slice(1);
  return v;
}

function carpetas_() {
  var ss = SpreadsheetApp.getActive();
  var padre = DriveApp.getFileById(ss.getId()).getParents();
  var base = padre.hasNext() ? padre.next() : DriveApp.getRootFolder();
  var sub = function (p, n) { var it = p.getFoldersByName(n); return it.hasNext() ? it.next() : p.createFolder(n); };
  var buzon = sub(base, 'buzon_claude');
  return { fotos: sub(base, 'fotos_caja'), buzon: buzon, procesados: sub(buzon, 'procesados') };
}

function hoja_(ss, nombre, headers) {
  var sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    var H = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    headers.forEach(function (h) { if (H.indexOf(h) < 0) { sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold'); } });
  }
  return sh;
}

function log_(origen, accion, detalle) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('BITACORA');
    if (sh) sh.appendRow([iso_(), origen, accion, String(detalle).slice(0, 500)]);
  } catch (err) {}
}

function nuevaClave_() {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 8; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s.slice(0, 4) + '-' + s.slice(4);
}

function slug_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 80);
}

function num_(v) {
  if (v === null || v === undefined || v === '') return null;
  var n = Number(String(v).replace(/[$,\s]/g, ''));
  return isFinite(n) ? n : null;
}

function col_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function iso_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
