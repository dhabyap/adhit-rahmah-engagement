const SHEET = "RSVP";
const TOKEN = "RA-ADMIT-2026";
const HOURS = 12;

function doPost(e) {
  var lock;
  try { lock = LockService.getScriptLock(); } catch (_) { lock = null; }
  if (lock) lock.enterBusy();
  try {
    let b;
  try { b = JSON.parse(e.postData.contents); } catch (_) { b = null; }
    if (!b || !b.nama) return json({ ok: false, error: "empty" });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET);
    if (!sh) sh = ss.insertSheet(SHEET);
    sh.getRange(1,1,1,6).setValues([[
      "Nama","Hadir?","Jumlah","Ucapan","HP","Waktu"
    ]]).setFontSize(10).setFontWeight("bold");
    const row = [
      String(b.nama||"").trim(),
      b.hadir === "yes",
      Math.max(1, Math.min(20, parseInt(b.jumlah||1,10))),
      String(b.ucapan||"").slice(0,500),
      String(b.hp||"").trim(),
      new Date()
    ];
    sh.appendRow(row);
    return json({ ok: true });
  } catch (err) {
    return json({ ok:false, error: String(err) });
  } finally { if (lock) lock.leave(); }
}

function doGet(e) {
  var lock;
  try { lock = LockService.getScriptLock(); } catch (_) { lock = null; }
  if (lock) lock.enterBusy();
  try {
    if ((e.parameter && e.parameter.key) !== TOKEN)
      return json({ error: "bad token" });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET) || ss.getSheets()[0];
    const last = sh.getLastRow();
    const out = [];
    if (last > 1) {
      sh.getRange(2,1,last-1,6).getDisplayValues().forEach(r =>
        out.push({
          nama: r[0],
          hadir: r[1] === "TRUE",
          jumlah: parseInt(r[2]||"0",10)||0,
          ucapan: r[3],
          hp: r[4],
          waktu: r[5]
        }));
    }
    return json({ ok: true, rows: out });
  } catch (err) {
    return json({ ok:false, error: String(err) });
  } finally { if (lock) lock.leave(); }
}

function json(o) {
  return ContentService.createTextOutput(
    JSON.stringify(o),
    ContentService.MimeType.JSON
  );
}

function setup() {
  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .insertSheet(SHEET)
    .setBackground("#F4F1EA")
    .getRange("A1:F1")
    .setValues([["Nama","Hadir?","Jumlah","Ucapan","HP","Waktu"]])
    .setFontWeight("bold");
  sh.getRange("A2:A1000").setNumberFormat("@");
  sh.getRange("F2:F1000").setNumberFormat("mm/dd/yyyy hh:mm:ss");
  sh.getRange("A2:A1000").setFontWeight("normal");
  return "Setup OK — sheet " + SHEET + " ready.";
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Setup")
    .addItem("Prepare sheet + styles", "setup")
    .addToUi();
}
