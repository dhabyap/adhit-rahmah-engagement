const SHEET = "RSVP";
const TOKEN = "RA-ADMIT-2026";

function doPost(e) {
  var body;
  try {
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.body) {
      body = JSON.parse(e.parameter.body);
    } else {
      return json({ ok: false, error: "empty" });
    }
  } catch (_) {
    return json({ ok: false, error: "bad json" });
  }

  if (!body || !body.nama) return json({ ok: false, error: "empty" });

  var lock = null;
  try { lock = LockService.getScriptLock(); } catch (_) { lock = null; }
  var gotLock = false;
  if (lock && lock.enterBusy) {
    try { gotLock = lock.enterBusy(); } catch (_) { gotLock = false; }
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) sh = ss.insertSheet(SHEET);
    var headers = ["Nama", "Hadir?", "Jumlah", "Ucapan", "HP", "Waktu"];
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontSize(10).setFontWeight("bold");

    var row = [
      String(body.nama || "").trim(),
      body.hadir === "yes",
      Math.max(1, Math.min(20, parseInt(body.jumlah || 1, 10) || 1)),
      String(body.ucapan || "").slice(0, 500),
      String(body.hp || "").trim(),
      new Date(),
      body.tampilPublik === true
    ];
    if (sh.getLastColumn() < 7) sh.getRange(1, 7).setValue("Tampil publik?").setFontWeight("bold");
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    if (gotLock && lock && lock.leave) lock.leave();
  }
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  // Public guestbook: expose only display name + greeting, never phone or guest count.
  if (p.public === "1") {
    try {
      var publicSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
      if (!publicSheet || publicSheet.getLastRow() < 2) return json({ ok: true, rows: [] });
      var publicValues = publicSheet.getRange(2, 1, publicSheet.getLastRow() - 1, Math.max(7, publicSheet.getLastColumn())).getDisplayValues();
      var publicRows = publicValues
        .filter(function(r) { return r[1] === "TRUE" && r[6] === "TRUE" && String(r[3] || "").trim(); })
        .map(function(r) { return { nama: r[0], ucapan: r[3] }; });
      for (var i = publicRows.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = publicRows[i]; publicRows[i] = publicRows[j]; publicRows[j] = temp;
      }
      return json({ ok: true, rows: publicRows.slice(0, 5) });
    } catch (err) {
      return json({ ok: false, rows: [] });
    }
  }

  if (p.key !== TOKEN) return json({ error: "bad token" });

  var lock = null;
  try { lock = LockService.getScriptLock(); } catch (_) { lock = null; }
  var gotLock = false;
  if (lock && lock.enterBusy) {
    try { gotLock = lock.enterBusy(); } catch (_) { gotLock = false; }
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET) || ss.getSheets()[0];
    var last = sh.getLastRow();
    var out = [];
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, 6).getDisplayValues();
      for (var i = 0; i < vals.length; i++) {
        var r = vals[i];
        out.push({
          nama: r[0],
          hadir: r[1] === "TRUE" || r[1] === "true",
          jumlah: parseInt(r[2] || "0", 10) || 0,
          ucapan: r[3] || "",
          hp: r[4] || "",
          waktu: r[5] || ""
        });
      }
    }
    return json({ ok: true, rows: out });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    if (gotLock && lock && lock.leave) lock.leave();
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o), ContentService.MimeType.JSON);
}

function setup() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET).setBackground("#F4F1EA");
  sh.getRange("A1:F1")
    .setValues([["Nama", "Hadir?", "Jumlah", "Ucapan", "HP", "Waktu"]])
    .setFontWeight("bold");
  sh.getRange("A2:A1000").setNumberFormat("@");
  sh.getRange("F2:F1000").setNumberFormat("mm/dd/yyyy hh:mm:ss");
  return "Setup OK — sheet " + SHEET + " ready.";
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Setup")
    .addItem("Prepare sheet + styles", "setup")
    .addToUi();
}
