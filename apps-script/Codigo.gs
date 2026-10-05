// Guarda las tareas marcadas como hechas en la carta Gantt del Proyecto Corazón.
// Va pegado en Extensiones → Apps Script de la planilla y se implementa como
// aplicación web (Ejecutar como: yo · Acceso: cualquier persona).

var HOJA = 'Tareas';

function hoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(['id', 'hecha', 'actualizado']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function leerEstado_() {
  var filas = hoja_().getDataRange().getValues().slice(1);
  var estado = {};
  filas.forEach(function (f) {
    if (f[0]) estado[f[0]] = f[1] === true || f[1] === 'TRUE';
  });
  return estado;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, state: leerEstado_() });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var id = String(body.id || '');
  if (!/^t\d{2}$/.test(id)) return json_({ ok: false, error: 'id inválido' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = hoja_();
    var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (f) { return f[0]; });
    var i = ids.indexOf(id);
    var fila = [id, body.done === true, new Date()];
    if (i === -1) sh.appendRow(fila);
    else sh.getRange(i + 1, 1, 1, 3).setValues([fila]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, state: leerEstado_() });
}
