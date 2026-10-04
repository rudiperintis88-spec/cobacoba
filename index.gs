/**
 * Backend sinkron Catatan & Tugas — Dashboard Piutang Unit
 * Dipasang di Google Sheet Anda (Extensions > Apps Script), lalu Deploy sebagai Web app.
 * Data disimpan di sheet bernama "Catatan" (dibuat otomatis).
 */
var SHEET_NAME = 'Catatan';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(['id', 'jenis', 'data_json', 'diubah', 'dihapus']);
    s.setFrozenRows(1);
  }
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Ambil semua catatan & tugas
function doGet(e) {
  var s = sheet_();
  var last = s.getLastRow();
  var items = [];
  if (last > 1) {
    var vals = s.getRange(2, 1, last - 1, 5).getValues();
    vals.forEach(function (r) {
      if (!r[0]) return;
      items.push({ id: String(r[0]), kind: r[1], json: r[2], deleted: String(r[4]) === '1' });
    });
  }
  return out_({ ok: true, items: items });
}

// Simpan perubahan: {ops:[{op:'upsert'|'delete', id, kind, json}]}
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    var ops = body.ops || [];
    var s = sheet_();
    var last = s.getLastRow();
    var rowOf = {};
    if (last > 1) {
      s.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { rowOf[String(r[0])] = i + 2; });
    }
    var now = new Date().getTime();
    ops.forEach(function (o) {
      var id = String(o.id);
      var row = rowOf[id];
      if (o.op === 'delete') {
        if (row) { s.getRange(row, 4, 1, 2).setValues([[now, 1]]); }
      } else {
        var vals = [[id, o.kind, o.json, now, 0]];
        if (row) { s.getRange(row, 1, 1, 5).setValues(vals); }
        else { s.appendRow(vals[0]); rowOf[id] = s.getLastRow(); }
      }
    });
    return out_({ ok: true, n: ops.length });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
