/**
 * Wochenrapport – Backend (Google Apps Script)
 * ------------------------------------------------------------------
 * Serviert das Web-Formular, sammelt alle Rapporte zentral in diesem
 * Google Sheet und erzeugt den CSV-Export im Messerli-Importformat
 * (Muster_Arbeitszeitimport-Vorlage.csv, Ticket BUS-SUP-27446):
 *
 *   Datum;Projektnummer;Mitarbeiter Nummer;Mitarbeiter;Bemerkung;Kostenart-Nummer;Dauer
 *   16.01.2026;26-26581.001;303;Markus Müller;;102;08:00
 *
 * Reiter:
 *   - "Rapporte"    : alle Eingaben (zentrale Ablage)
 *   - "Projekte"    : Spalte A Projektnummer (exakt wie in MesserliPROJEKT!), Spalte B Bezeichnung
 *   - "Mitarbeiter" : Spalte A Mitarbeiter-Nummer (exakt!), Spalte B Name
 *   - "Kostenarten" : Spalte A Typ (Arbeit / Absenzgründe), Spalte B Kostenart-Nummer (exakt!)
 */

var SHEET_RAPPORTE    = 'Rapporte';
var SHEET_PROJEKTE    = 'Projekte';
var SHEET_MITARBEITER = 'Mitarbeiter';
var SHEET_KOSTEN     = 'Kostenarten';
var SOLL_STUNDEN      = 40;

/** Absenzgründe – identisch zur Liste im Formular. */
var ABSENZEN = ['Büro', 'Ferien', 'Krank', 'Unfall', 'Feiertag', 'Schule', 'Militär/ZS', 'Weiterbildung'];

/** Menü im Sheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wochenrapport')
    .addItem('Reiter einrichten', 'ensureSheets_')
    .addItem('Messerli-Export erstellen …', 'exportMesserli')
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
    ss.insertSheet(SHEET_RAPPORTE).appendRow(
      ['Zeitstempel', 'Mitarbeiter Nr.', 'Mitarbeiter', 'Woche vom', 'Rapport Nr.', 'Datum', 'Auftrag Nr.', 'Stunden']);
  }
  if (!ss.getSheetByName(SHEET_PROJEKTE)) {
    ss.insertSheet(SHEET_PROJEKTE).appendRow(['Projektnummer (exakt wie Messerli)', 'Bezeichnung']);
  }
  if (!ss.getSheetByName(SHEET_MITARBEITER)) {
    ss.insertSheet(SHEET_MITARBEITER).appendRow(['Mitarbeiter Nr. (exakt wie Messerli)', 'Name']);
  }
  if (!ss.getSheetByName(SHEET_KOSTEN)) {
    var ks = ss.insertSheet(SHEET_KOSTEN);
    ks.appendRow(['Typ', 'Kostenart-Nummer (exakt wie Messerli)']);
    ks.appendRow(['Arbeit', '']);          // produktive Projektstunden
    ABSENZEN.forEach(function (a) { ks.appendRow([a, '']); });
  }
}

/** Stammdaten fürs Formular: Namen (mit Nr.), Projekte, Sollzeit. */
function getStammdaten() {
  ensureSheets_();
  var ss = ss_();

  var projects = [];
  var proj = ss.getSheetByName(SHEET_PROJEKTE).getDataRange().getValues();
  for (var i = 1; i < proj.length; i++) {
    var nr = trim_(proj[i][0]);
    if (nr) projects.push({ nr: nr, label: trim_(proj[i][1]) });
  }

  var names = [];
  var mit = ss.getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  for (var j = 1; j < mit.length; j++) {
    var mnr = trim_(mit[j][0]), name = trim_(mit[j][1]);
    if (name) names.push({ nr: mnr, name: name });
  }

  return { projects: projects, names: names, soll: SOLL_STUNDEN };
}

/**
 * Speichert einen Wochenrapport zentral. payload = {
 *   ma, maNr, von, rapnr, entries: [ { datum, auftrag, stunden }, ... ]
 * }
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
      return [now, String(payload.maNr || ''), String(payload.ma), String(payload.von || ''),
              String(payload.rapnr || ''), String(e.datum || ''), String(e.auftrag), Number(e.stunden)];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/* ================= Messerli-Export ================= */

/**
 * Fragt einen Datumsbereich ab, baut die CSV im Messerli-Format und
 * legt sie in Google Drive ab (Link wird angezeigt).
 */
function exportMesserli() {
  var ui = SpreadsheetApp.getUi();
  ensureSheets_();

  var von = askDate_(ui, 'Export von (TT.MM.JJJJ)'); if (von === null) return;
  var bis = askDate_(ui, 'Export bis (TT.MM.JJJJ)'); if (bis === null) return;

  var kosten = readKostenarten_();
  var mitByName = readMitarbeiter_();

  var data = ss_().getSheetByName(SHEET_RAPPORTE).getDataRange().getValues();
  var lines = ['Datum;Projektnummer;Mitarbeiter Nummer;Mitarbeiter;Bemerkung;Kostenart-Nummer;Dauer'];
  var warn = {}, count = 0;

  for (var i = 1; i < data.length; i++) {
    var maNr = trim_(data[i][1]), ma = trim_(data[i][2]);
    var datum = parseChDate_(trim_(data[i][5]));
    var auftrag = trim_(data[i][6]);
    var stunden = Number(data[i][7]);
    if (!datum || !auftrag || !(stunden > 0)) continue;
    if (datum < von || datum > bis) continue;

    if (!maNr) maNr = mitByName[ma.toLowerCase()] || '';
    if (!maNr) warn['Mitarbeiter-Nr. fehlt: ' + ma] = 1;

    var istAbsenz = kosten.absenz.hasOwnProperty(auftrag);
    var projektnummer = istAbsenz ? '' : auftrag;
    var bemerkung = istAbsenz ? auftrag : '';
    var kostenart = istAbsenz ? kosten.absenz[auftrag] : kosten.arbeit;
    if (!kostenart) warn['Kostenart-Nummer fehlt für: ' + (istAbsenz ? auftrag : 'Arbeit')] = 1;

    lines.push([fmtChDate_(datum), projektnummer, maNr, ma, bemerkung, kostenart, toHHMM_(stunden)]
      .map(csvCell_).join(';'));
    count++;
  }

  if (!count) { ui.alert('Keine Buchungen im gewählten Zeitraum gefunden.'); return; }

  var name = 'Messerli-Import_' + fmtChDate_(von).replace(/\./g, '-') + '_bis_' + fmtChDate_(bis).replace(/\./g, '-') + '.csv';
  var file = DriveApp.createFile(name, '﻿' + lines.join('\r\n'), 'text/csv');

  var msg = count + ' Buchungszeilen exportiert.\n\nDatei in Google Drive:\n' + file.getUrl();
  var warnKeys = Object.keys(warn);
  if (warnKeys.length) msg += '\n\n⚠ Bitte prüfen:\n- ' + warnKeys.join('\n- ');
  ui.alert('Messerli-Export erstellt', msg, ui.ButtonSet.OK);
}

/* ---------------- Hilfsfunktionen ---------------- */

function readKostenarten_() {
  var out = { arbeit: '', absenz: {} };
  var v = ss_().getSheetByName(SHEET_KOSTEN).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var typ = trim_(v[i][0]), nr = trim_(v[i][1]);
    if (!typ) continue;
    if (typ.toLowerCase() === 'arbeit') out.arbeit = nr;
    else out.absenz[typ] = nr;
  }
  return out;
}

function readMitarbeiter_() {
  var map = {};
  var v = ss_().getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var nr = trim_(v[i][0]), name = trim_(v[i][1]);
    if (name && nr) map[name.toLowerCase()] = nr;
  }
  return map;
}

function askDate_(ui, promptText) {
  var res = ui.prompt('Messerli-Export', promptText, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  var d = parseChDate_(res.getResponseText().trim());
  if (!d) { ui.alert('Ungültiges Datum: "' + res.getResponseText() + '" – bitte als TT.MM.JJJJ eingeben.'); return null; }
  return d;
}

/** "16.01.2026" -> Date (12:00 lokal, damit Zeitzonen keine Tagesverschiebung verursachen). */
function parseChDate_(s) {
  var m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(String(s || '').trim());
  if (!m) return null;
  var d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

function fmtChDate_(d) {
  return ('0' + d.getDate()).slice(-2) + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear();
}

/** Dezimalstunden -> "HH:MM" (8.5 -> "08:30"). Minuten werden auf ganze gerundet. */
function toHHMM_(hours) {
  var totalMin = Math.round(Number(hours) * 60);
  var h = Math.floor(totalMin / 60), m = totalMin % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

function csvCell_(v) {
  v = (v == null ? '' : String(v));
  return /[";\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}

function trim_(v) { return (v == null ? '' : String(v)).trim(); }
