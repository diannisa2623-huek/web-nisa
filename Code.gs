/**
 * Loka Kota – Backend Google Apps Script
 * Tempel kode ini di Extensions > Apps Script pada spreadsheet kamu.
 */
var SHEET_NAME = "Votes";
var CITIES = ["Yogyakarta", "Bandung", "Jakarta", "Malang", "Surabaya", "Denpasar", "Medan", "Makassar"];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Waktu", "Nama", "Kota", "Alasan"]);
    sh.getRange("A1:D1").setFontWeight("bold").setBackground("#cfeafb");
    sh.setFrozenRows(1);
  }
  return sh;
}

// Mencegah teks diperlakukan sebagai rumus di spreadsheet
function safe_(v) {
  v = String(v == null ? "" : v).replace(/[\r\n]+/g, " ").trim();
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    var counts = {};
    CITIES.forEach(function (c) { counts[c] = 0; });
    var recent = [], total = 0;

    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 4).getValues();
      rows.forEach(function (r) {
        var city = String(r[2]);
        if (counts.hasOwnProperty(city)) { counts[city]++; total++; }
      });
      rows.slice(-6).reverse().forEach(function (r) {
        recent.push({ name: String(r[1]), city: String(r[2]), reason: String(r[3]) });
      });
    }
    return json_({ ok: true, counts: counts, total: total, recent: recent });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var data = JSON.parse(e.postData.contents);
    var city = String(data.city || "");
    if (CITIES.indexOf(city) === -1) {
      return json_({ ok: false, error: "Kota tidak valid" });
    }
    getSheet_().appendRow([
      new Date(),
      safe_(data.name).substring(0, 40),
      city,
      safe_(data.reason).substring(0, 200)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
