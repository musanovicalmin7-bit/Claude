/**
 * Wochenrapport – Backend (Google Apps Script)
 * ------------------------------------------------------------------
 * Serviert das Web-Formular und sammelt alle Rapporte zentral in
 * diesem Google Sheet. Drei Reiter werden verwendet:
 *   - "Rapporte"    : hier laufen alle Eingaben zusammen (Auswertung/Export)
 *   - "Projekte"    : Auftragsnummern (Spalte A) + Bezeichnung (Spalte B)
 *   - "Mitarbeiter" : Namen (Spalte A)
 * Projekte und Mitarbeiter pflegst du per Copy-Paste aus Messerli.
 */

var SHEET_RAPPORTE    = 'Rapporte';
var SHEET_PROJEKTE    = 'Projekte';
var SHEET_MITARBEITER = 'Mitarbeiter';
var SOLL_STUNDEN      = 40;

/** Menü zum Einrichten, sobald das Sheet geöffnet wird. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wochenrapport')
    .addItem('Reiter einrichten', 'ensureSheets_')
    .addToUi();
}

/** Liefert das Web-Formular aus. */
function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Wochenrapport')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ss_() { return SpreadsheetApp.getActive(); }

/** Legt fehlende Reiter mit Kopfzeilen an (idempotent). */
function ensureSheets_() {
  var ss = ss_();
  if (!ss.getSheetByName(SHEET_RAPPORTE)) {
    ss.insertSheet(SHEET_RAPPORTE)
      .appendRow(['Zeitstempel', 'Mitarbeiter', 'Woche vom', 'Rapport Nr.', 'Datum', 'Auftrag Nr.', 'Stunden']);
  }
  if (!ss.getSheetByName(SHEET_PROJEKTE)) {
    ss.insertSheet(SHEET_PROJEKTE).appendRow(['Auftrag Nr.', 'Bezeichnung']);
  }
  if (!ss.getSheetByName(SHEET_MITARBEITER)) {
    ss.insertSheet(SHEET_MITARBEITER).appendRow(['Name']);
  }
}

/** Wird vom Formular beim Laden aufgerufen: liefert Namen + Projekte + Sollzeit. */
function getStammdaten() {
  ensureSheets_();
  var ss = ss_();

  var projects = [];
  var proj = ss.getSheetByName(SHEET_PROJEKTE).getDataRange().getValues();
  for (var i = 1; i < proj.length; i++) {          // Zeile 0 = Kopfzeile
    var nr = trim_(proj[i][0]);
    var label = trim_(proj[i][1]);
    if (nr) projects.push({ nr: nr, label: label });
  }

  var names = [];
  var mit = ss.getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  for (var j = 1; j < mit.length; j++) {
    var n = trim_(mit[j][0]);
    if (n) names.push(n);
  }

  return { projects: projects, names: names, soll: SOLL_STUNDEN };
}

/**
 * Speichert einen Wochenrapport zentral. payload = {
 *   ma, von, rapnr, entries: [ { datum, auftrag, stunden }, ... ]
 * }
 * Ein Lock verhindert Konflikte, wenn viele gleichzeitig absenden.
 */
function submitRapport(payload) {
  if (!payload || !payload.ma) throw new Error('Bitte zuerst den Namen wählen.');
  var entries = (payload.entries || []).filter(function (e) {
    return e && e.auftrag && Number(e.stunden) > 0;
  });
  if (!entries.length) throw new Error('Keine Stunden erfasst.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    var sheet = ss_().getSheetByName(SHEET_RAPPORTE);
    var now = new Date();
    var rows = entries.map(function (e) {
      return [now, String(payload.ma), String(payload.von || ''), String(payload.rapnr || ''),
              String(e.datum || ''), String(e.auftrag), Number(e.stunden)];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function trim_(v) { return (v == null ? '' : String(v)).trim(); }
