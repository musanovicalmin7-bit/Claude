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

/**
 * Absenzarten mit Kostenart-Nummern aus MesserliPROJEKT (Stand 08/2026).
 * Wird beim Einrichten in den Reiter "Kostenarten" geschrieben und kann
 * dort jederzeit angepasst werden – das Sheet ist die führende Quelle.
 */
var ABSENZEN = [
  { nr: '001', label: 'Unfall' },
  { nr: '002', label: 'Unfall 20% Abzug' },
  { nr: '003', label: 'Krank 20% Abzug' },
  { nr: '004', label: 'Krank' },
  { nr: '005', label: 'Militär / Zivilschutz' },
  { nr: '006', label: 'Militär / Zivilschutz 50% Abzug' },
  { nr: '007', label: 'Ferien Bezug' },
  { nr: '008', label: 'Ferien Gutschrift' },
  { nr: '009', label: 'Schule' },
  { nr: '010', label: 'Weiterbildung' },
  { nr: '011', label: 'Heirat' },
  { nr: '012', label: 'Geburt' },
  { nr: '013', label: 'Todesfall' },
  { nr: '014', label: 'Öffentliches Amt' },
  { nr: '015', label: 'Bezahlte Absenzen' },
  { nr: '016', label: 'Unbezahlte Absenzen' },
  { nr: '017', label: 'Diverse Absenzen' }
];

/** Menü im Sheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wochenrapport')
    .addItem('Reiter einrichten', 'ensureSheets_')
    .addItem('Kontrolle aktualisieren', 'updateKontrolle')
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
  var rapHeader = ['Zeitstempel', 'Mitarbeiter Nr.', 'Mitarbeiter', 'Woche vom', 'Rapport Nr.', 'Datum', 'Auftrag Nr.', 'Stunden', 'Rapport-PDF (signiert)'];
  var rap = ss.getSheetByName(SHEET_RAPPORTE);
  if (!rap) rap = ss.insertSheet(SHEET_RAPPORTE);
  rap.getRange(1, 1, 1, rapHeader.length).setValues([rapHeader]).setFontWeight('bold');
  if (!ss.getSheetByName(SHEET_PROJEKTE)) {
    ss.insertSheet(SHEET_PROJEKTE).appendRow(['Projektnummer (exakt wie Messerli)', 'Bezeichnung']);
  }
  if (!ss.getSheetByName(SHEET_MITARBEITER)) {
    ss.insertSheet(SHEET_MITARBEITER).appendRow(['Mitarbeiter Nr. (exakt wie Messerli)', 'Name']);
  }
  if (!ss.getSheetByName(SHEET_KOSTEN)) {
    var ks = ss.insertSheet(SHEET_KOSTEN);
    ks.appendRow(['Typ', 'Kostenart-Nummer (exakt wie Messerli)']);
    ks.appendRow(['Arbeit', '']);          // produktive Projektstunden – Nummer noch eintragen!
    ABSENZEN.forEach(function (a) { ks.appendRow([a.label, a.nr]); });
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

  // Absenzarten aus dem Kostenarten-Reiter (ohne "Arbeit") – so zeigt das
  // Formular exakt die Liste, die auch der Export verwendet.
  var absenzen = [];
  var kost = ss.getSheetByName(SHEET_KOSTEN).getDataRange().getValues();
  for (var k = 1; k < kost.length; k++) {
    var typ = trim_(kost[k][0]);
    if (typ && typ.toLowerCase() !== 'arbeit') absenzen.push(typ);
  }

  return { projects: projects, names: names, absenzen: absenzen, soll: SOLL_STUNDEN };
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

  // GAV-konformes, unterschriebenes Wochenrapport-PDF erzeugen und in Drive archivieren.
  var pdfUrl = '';
  try {
    pdfUrl = createRapportPdf_(payload, entries);
  } catch (err) {
    pdfUrl = 'PDF-Fehler: ' + err;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets_();
    var sheet = ss_().getSheetByName(SHEET_RAPPORTE);
    var now = new Date();
    var rows = entries.map(function (e) {
      return [now, String(payload.maNr || ''), String(payload.ma), String(payload.von || ''),
              String(payload.rapnr || ''), String(e.datum || ''), String(e.auftrag), Number(e.stunden),
              pdfUrl];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Baut das Wochenrapport-PDF (analog Papierformular EIT.swiss):
 * Stundentabelle pro Auftrag, Protokoll Arbeitszeiten (Beginn/Pause/Ende, GAV
 * Art. 20/28 bzw. ArG) und die Unterschrift des Mitarbeiters. Ablage im
 * Drive-Ordner "Wochenrapport PDFs", Rückgabe: Datei-URL.
 */
function createRapportPdf_(payload, entries) {
  var monday = parseAnyDate_(payload.von) || mondayOf_(new Date());
  monday = mondayOf_(monday);
  var days = [], DW = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
  for (var i = 0; i < 7; i++) {
    var d = new Date(monday); d.setDate(d.getDate() + i); days.push(d);
  }
  var kw = isoWeek_(monday);

  // Stunden je Auftrag x Tag gruppieren
  var byAuftrag = {}, order = [], colTot = [0, 0, 0, 0, 0, 0, 0], tot = 0;
  entries.forEach(function (e) {
    var idx = -1;
    for (var i = 0; i < 7; i++) { if (fmtChDate_(days[i]) === String(e.datum)) { idx = i; break; } }
    if (idx < 0) return;
    var key = String(e.auftrag);
    if (!byAuftrag[key]) { byAuftrag[key] = [0, 0, 0, 0, 0, 0, 0]; order.push(key); }
    byAuftrag[key][idx] += Number(e.stunden);
    colTot[idx] += Number(e.stunden); tot += Number(e.stunden);
  });

  function td(v, style) { return '<td style="border:1px solid #999;padding:4px 6px;' + (style || '') + '">' + v + '</td>'; }
  function num(v) { return v ? String(Math.round(v * 100) / 100).replace('.', ',') : ''; }

  var h = '<html><head><meta charset="utf-8"></head><body style="font-family:Helvetica,Arial,sans-serif;font-size:11px;color:#000">' +
    '<h1 style="font-size:20px;border-bottom:3px solid #2E5E8C;padding-bottom:6px;margin-bottom:4px">Wochenrapport</h1>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:10px"><tr>' +
    '<td style="padding:2px 0"><b>Mitarbeiter:</b> ' + esc_(payload.ma) + (payload.maNr ? ' (Nr. ' + esc_(payload.maNr) + ')' : '') + '</td>' +
    '<td style="padding:2px 0"><b>Woche:</b> ' + fmtChDate_(days[0]) + ' – ' + fmtChDate_(days[6]) + '</td>' +
    '<td style="padding:2px 0"><b>KW:</b> ' + kw + '</td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse"><tr style="background:#EDF3F8">' +
    td('<b>Auftrag Nr. / Absenz</b>');
  days.forEach(function (d, i) { h += td('<b>' + DW[i] + '</b><br>' + ('0' + d.getDate()).slice(-2) + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.', 'text-align:center'); });
  h += td('<b>Total</b>', 'text-align:right') + '</tr>';
  order.forEach(function (key) {
    var row = byAuftrag[key], rt = 0;
    h += '<tr>' + td(esc_(key));
    row.forEach(function (v) { rt += v; h += td(num(v), 'text-align:center'); });
    h += td('<b>' + num(rt) + '</b>', 'text-align:right') + '</tr>';
  });
  h += '<tr style="background:#EDF3F8">' + td('<b>Total Stunden</b>');
  colTot.forEach(function (v) { h += td('<b>' + num(v) + '</b>', 'text-align:center'); });
  h += td('<b>' + num(tot) + '</b>', 'text-align:right') + '</tr></table>' +
    '<h2 style="font-size:14px;margin:16px 0 6px">Protokoll Arbeitszeiten</h2>' +
    '<table style="width:60%;border-collapse:collapse"><tr style="background:#EDF3F8">' +
    td('<b>Tag</b>') + td('<b>Beginn</b>', 'text-align:center') + td('<b>Pause</b>', 'text-align:center') +
    td('<b>Ende</b>', 'text-align:center') + td('<b>Stunden</b>', 'text-align:right') + '</tr>';
  (payload.protokoll || []).forEach(function (p) {
    var pause = (p.pauseVon && p.pauseBis) ? (p.pauseVon + ' – ' + p.pauseBis) : '–';
    h += '<tr>' + td(esc_(p.tag) + ' ' + esc_(p.datum)) + td(esc_(p.start), 'text-align:center') +
         td(pause, 'text-align:center') + td(esc_(p.ende), 'text-align:center') +
         td(num(p.stunden), 'text-align:right') + '</tr>';
  });
  h += '</table>';

  var m = /^data:image\/(png|jpeg);base64,(.+)$/.exec(String(payload.unterschrift || ''));
  h += '<table style="width:100%;margin-top:26px"><tr>' +
    '<td style="width:50%;vertical-align:bottom;padding:2px 0">Erfasst am: ' + fmtChDate_(new Date()) + '</td>' +
    '<td style="width:50%;vertical-align:bottom">' +
    (m ? '<img src="' + payload.unterschrift + '" style="height:60px"><br>' : '<br><br>') +
    '<span style="border-top:1px solid #000;display:inline-block;padding-top:3px;min-width:220px">Unterschrift Mitarbeiter</span></td></tr></table>' +
    '</body></html>';

  var name = 'Wochenrapport_' + days[0].getFullYear() + '_KW' + ('0' + kw).slice(-2) + '_' +
             String(payload.ma).replace(/[^\wÀ-ÿ]+/g, '_') + '.pdf';
  var pdf = Utilities.newBlob(h, 'text/html', 'r.html').getAs('application/pdf').setName(name);
  var folderName = 'Wochenrapport PDFs';
  var it = DriveApp.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  return folder.createFile(pdf).getUrl();
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/* ================= Kontroll-Übersicht ================= */

var SHEET_KONTROLLE = 'Kontrolle';

/**
 * Baut den Reiter "Kontrolle": Zeilen = alle Mitarbeiter, Spalten = Kalender-
 * wochen (von der ersten eingereichten Woche bis zur aktuellen). Zellwert =
 * eingereichte Stunden. Grün >= Sollzeit, Gelb = teilweise, Rot = fehlt.
 */
function updateKontrolle() {
  ensureSheets_();
  var ss = ss_();

  var employees = [];
  var mit = ss.getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  for (var i = 1; i < mit.length; i++) {
    var nr = trim_(mit[i][0]), name = trim_(mit[i][1]);
    if (name) employees.push({ nr: nr, name: name });
  }
  if (!employees.length) {
    SpreadsheetApp.getUi().alert('Der Reiter "Mitarbeiter" ist leer – bitte zuerst die Namen eintragen.');
    return;
  }

  // Stunden pro Mitarbeiter und Woche (Montag) aufsummieren
  var data = ss.getSheetByName(SHEET_RAPPORTE).getDataRange().getValues();
  var byWeek = {};                       // nameLower -> { mondayIso -> stunden }
  var minMonday = null;
  for (var r = 1; r < data.length; r++) {
    var name = trim_(data[r][2]);
    var monday = parseAnyDate_(data[r][3]) || parseAnyDate_(data[r][5]);
    var stunden = Number(data[r][7]);
    if (!name || !monday || !(stunden > 0)) continue;
    monday = mondayOf_(monday);
    var iso = isoDate_(monday);
    var key = name.toLowerCase();
    if (!byWeek[key]) byWeek[key] = {};
    byWeek[key][iso] = (byWeek[key][iso] || 0) + stunden;
    if (!minMonday || monday < minMonday) minMonday = monday;
  }

  var thisMonday = mondayOf_(new Date());
  if (!minMonday) minMonday = thisMonday;
  var mondays = [];
  for (var d = new Date(minMonday); d <= thisMonday; d.setDate(d.getDate() + 7)) {
    mondays.push(new Date(d));
  }

  // Tabelle aufbauen
  var header = ['Nr.', 'Mitarbeiter'];
  mondays.forEach(function (m) {
    header.push('KW ' + isoWeek_(m) + '\n' + ('0' + m.getDate()).slice(-2) + '.' + ('0' + (m.getMonth() + 1)).slice(-2) + '.');
  });
  var rows = [header], colors = [header.map(function () { return '#EDF3F8'; })];
  employees.forEach(function (e) {
    var row = [e.nr, e.name], color = ['', ''];
    var wk = byWeek[e.name.toLowerCase()] || {};
    mondays.forEach(function (m) {
      var h = wk[isoDate_(m)] || 0;
      row.push(h ? h : '');
      color.push(h >= SOLL_STUNDEN ? '#D9EAD3' : (h > 0 ? '#FFF2CC' : '#F4CCCC'));
    });
    rows.push(row); colors.push(color);
  });

  var sheet = ss.getSheetByName(SHEET_KONTROLLE) || ss.insertSheet(SHEET_KONTROLLE);
  sheet.clear();
  sheet.getRange(1, 1, rows.length, header.length).setValues(rows).setBackgrounds(colors);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setWrap(true);
  sheet.setFrozenRows(1); sheet.setFrozenColumns(2);
  sheet.autoResizeColumns(1, 2);
  for (var c = 3; c <= header.length; c++) sheet.setColumnWidth(c, 62);
  ss.setActiveSheet(sheet);
}

/** ISO-Kalenderwoche (KW) einer Date. */
function isoWeek_(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dayNum + 3);
  var firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  var firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);
  return 1 + Math.round((d - firstThursday) / (7 * 24 * 3600 * 1000));
}

function mondayOf_(d) {
  var x = new Date(d); var day = (x.getDay() + 6) % 7;
  x.setDate(x.getDate() - day); x.setHours(12, 0, 0, 0); return x;
}

function isoDate_(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

/** Akzeptiert "2026-08-03", "03.08.2026" oder ein Date-Objekt. */
function parseAnyDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v || '').trim();
  var iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 12, 0, 0);
  return parseChDate_(s);
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
