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
 *   - "Anträge"     : Ferienanträge und Absenzmeldungen mit Status (offen/bewilligt/abgelehnt/zurückgezogen)
 */

var SHEET_RAPPORTE    = 'Rapporte';
var SHEET_PROJEKTE    = 'Projekte';
var SHEET_MITARBEITER = 'Mitarbeiter';
var SHEET_KOSTEN     = 'Kostenarten';
var SHEET_ENTWURF    = 'Entwürfe';      // Zwischenstände der App (pro Mitarbeiter und Woche), ausgeblendet
var SHEET_ANTRAG     = 'Anträge';
var SOLL_STUNDEN      = 40;

/**
 * Empfänger, der JEDEN abgesendeten Wochenrapport als PDF-Anhang erhält
 * (z. B. Büro/Administration). Leer lassen = kein automatischer Versand.
 */
var RAPPORT_EMPFAENGER = 'elektroscherzingerag@gmail.com';

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
  { nr: '017', label: 'Diverse Absenzen' },
  // Später ergänzt: wird auch in einen bestehenden Kostenarten-Reiter nachgetragen.
  { nr: 'K-ÜS', label: 'Kompensation Überstunden (frei)', neu: true }
];

/**
 * Test: schickt eine Probemail an RAPPORT_EMPFAENGER.
 * Im Editor ausführen – beim ersten Mal fragt Google nach der
 * E-Mail-Berechtigung ("E-Mails in deinem Namen senden") → zulassen.
 */
function testMail() {
  if (!RAPPORT_EMPFAENGER) throw new Error('RAPPORT_EMPFAENGER ist leer.');
  MailApp.sendEmail(RAPPORT_EMPFAENGER, 'Test Wochenrapport',
    'Testmail vom Wochenrapport-Script. Wenn diese Mail ankommt, funktioniert der Versand.');
}

/**
 * Test: schreibt einen Probe-Entwurf ins Blatt "Entwürfe", liest ihn zurück und löscht ihn wieder.
 * Im Editor ausführen → im Protokoll muss "Entwurf-Test OK" stehen. Beim ersten Mal Berechtigung zulassen.
 */
function testEntwurf() {
  var name = '_Test Entwurf_', week = '2026-01-05';
  saveEntwurf('0', name, week, '{"cards":[{"auftrag":"TEST","h":[1,0,0,0,0,0,0]}]}');
  var map = loadEntwuerfe(name);
  saveEntwurf('0', name, week, null);
  if (!map[week] || map[week].json.indexOf('TEST') < 0) throw new Error('Entwurf-Test FEHLGESCHLAGEN: ' + JSON.stringify(map));
  Logger.log('Entwurf-Test OK – Blatt "' + SHEET_ENTWURF + '" funktioniert.');
  return 'OK';
}

/** Menü im Sheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wochenrapport')
    .addItem('Reiter einrichten', 'ensureSheets_')
    .addItem('Kontrolle aktualisieren', 'updateKontrolle')
    .addItem('Messerli-Export erstellen …', 'exportMesserli')
    .addSeparator()
    .addItem('Erinnerung offene Anträge einrichten (täglich)', 'setupErinnerung')
    .addItem('Test: Antrag-Mail an GL senden', 'testAntragMail')
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
  var mitHeader = ['Mitarbeiter Nr. (exakt wie Messerli)', 'Name', 'E-Mail (optional, für PDF-Zustellung)',
                   'Kostenart Arbeit (je nach Funktion: Service/Installateur/Monteur/Lehrling …)',
                   'PIN (optional, z. B. 4-stellig – schützt vor Erfassung unter fremdem Namen)',
                   'Rolle (GL = darf Anträge bewilligen, braucht PIN)'];
  var mit = ss.getSheetByName(SHEET_MITARBEITER);
  if (!mit) mit = ss.insertSheet(SHEET_MITARBEITER);
  mit.getRange(1, 1, 1, mitHeader.length).setValues([mitHeader]).setFontWeight('bold');
  if (!ss.getSheetByName(SHEET_KOSTEN)) {
    var ks = ss.insertSheet(SHEET_KOSTEN);
    ks.appendRow(['Typ', 'Kostenart-Nummer (exakt wie Messerli)']);
    ks.appendRow(['Arbeit', '']);          // produktive Projektstunden – Nummer noch eintragen!
    ABSENZEN.forEach(function (a) { ks.appendRow([a.label, a.nr]); });
  } else {
    nachtragKostenarten_(ss.getSheetByName(SHEET_KOSTEN));
  }
}

/** Trägt neu dazugekommene Absenzarten nach, falls Typ und Nummer noch fehlen. */
function nachtragKostenarten_(ks) {
  var v = ks.getDataRange().getValues();
  var typen = {}, nummern = {};
  for (var i = 1; i < v.length; i++) {
    typen[trim_(v[i][0]).toLowerCase()] = true;
    nummern[trim_(v[i][1]).toLowerCase()] = true;
  }
  ABSENZEN.forEach(function (a) {
    if (a.neu && !typen[a.label.toLowerCase()] && !nummern[a.nr.toLowerCase()]) ks.appendRow([a.label, a.nr]);
  });
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

  // pin: nur JA/NEIN – der PIN selbst verlässt den Server nie.
  var names = [];
  var mit = ss.getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  for (var j = 1; j < mit.length; j++) {
    var mnr = trim_(mit[j][0]), name = trim_(mit[j][1]);
    if (name) names.push({ nr: mnr, name: name, pin: !!trim_(mit[j][4]) });
    // Rolle wird bewusst nicht mitgeschickt – GL-Rechte prüft der Server erst nach dem PIN.
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
/**
 * Prüft den PIN eines Mitarbeiters (Spalte E). Ohne hinterlegten PIN immer ok.
 * Nach 5 Fehlversuchen ist der Name für 15 Minuten gesperrt.
 * gl = true nur, wenn Rolle "GL" und ein PIN hinterlegt und korrekt ist.
 */
function checkPin(name, pin) {
  var key = 'pinfail_' + String(name || '').trim().toLowerCase();
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get(key) || 0);
  if (fails >= 5) return { ok: false, locked: true };
  var rec = readMitarbeiter_()[String(name || '').trim().toLowerCase()] || {};
  var ok = !rec.pin || rec.pin === String(pin || '').trim();
  if (!ok) cache.put(key, String(fails + 1), 15 * 60);
  return { ok: ok, gl: ok && !!rec.pin && /^gl$/i.test(rec.rolle || '') };
}
function requirePin_(name, pin) {
  var r = checkPin(name, pin);
  if (r.locked) throw new Error('Zu viele falsche PIN-Eingaben – bitte 15 Minuten warten.');
  if (!r.ok) throw new Error('Falscher PIN.');
  return r;
}

function submitRapport(payload) {
  if (!payload || !payload.ma) throw new Error('Bitte zuerst den Namen wählen.');
  requirePin_(payload.ma, payload.pin);
  var entries = (payload.entries || []).filter(function (e) {
    return e && e.auftrag && Number(e.stunden) > 0;
  });
  if (!entries.length) throw new Error('Keine Stunden erfasst.');

  // GAV-konformes, unterschriebenes Wochenrapport-PDF erzeugen und in Drive archivieren.
  var pdfUrl = '', pdfFile = null;
  try {
    pdfFile = createRapportPdf_(payload, entries);
    // Mitarbeiter darf sein eigenes PDF über den Link ansehen (Link ist nicht erratbar).
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    pdfUrl = pdfFile.getUrl();
  } catch (err) {
    pdfUrl = 'PDF-Fehler: ' + err;
  }

  // E-Mail-Zustellung des PDFs:
  //  1. an die Administration (RAPPORT_EMPFAENGER), wenn konfiguriert
  //  2. an den Mitarbeiter, wenn im Reiter "Mitarbeiter" (Spalte C) eine Adresse steht
  var mailInfo = '';
  if (pdfFile) {
    var subject = pdfFile.getName().replace('.pdf', '').replace(/_/g, ' ');
    if (RAPPORT_EMPFAENGER) {
      try {
        MailApp.sendEmail({
          to: RAPPORT_EMPFAENGER,
          subject: subject,
          body: 'Neuer Wochenrapport eingegangen von ' + payload.ma +
                (payload.maNr ? ' (Nr. ' + payload.maNr + ')' : '') + '.\n\n' +
                'PDF im Anhang und im Drive-Ordner "Wochenrapport PDFs".\n' +
                'Diese Nachricht wurde automatisch erstellt.',
          attachments: [pdfFile.getBlob()]
        });
      } catch (err) {}
    }
    try {
      var email = emailForName_(payload.ma);
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: 'Hallo ' + payload.ma + '\n\nIm Anhang dein unterschriebener Wochenrapport.\n\n' +
                'Diese Nachricht wurde automatisch erstellt.',
          attachments: [pdfFile.getBlob()]
        });
        mailInfo = email;
      }
    } catch (err) { mailInfo = ''; }
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
    deleteEntwurf_(payload.ma, payload.von);   // Zwischenstand ist mit der Abgabe erledigt
    return { ok: true, count: rows.length,
             pdfUrl: (pdfFile ? pdfUrl : ''), mailedTo: mailInfo };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------
 * Entwürfe: Zwischenstand einer angefangenen Woche wird zentral gesichert.
 * Grund: Der Browser-Speicher der App ist im eingebetteten Google-Fenster
 * nicht zuverlässig (wechselnde Ursprünge, blockierte Cookies) – nach dem
 * Schliessen der App war ein angefangener Rapport weg.
 * ------------------------------------------------------------------ */
function entwurfSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SHEET_ENTWURF);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ENTWURF);
    sh.appendRow(['Mitarbeiter Nr.', 'Mitarbeiter', 'Woche vom (ISO)', 'Entwurf (JSON)', 'Zuletzt gespeichert']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
/** Wochenwert aus dem Blatt immer als "YYYY-MM-DD" (Sheets wandelt solche Texte gern in ein Datum um). */
function entwurfWeek_(w) {
  if (w instanceof Date) return Utilities.formatDate(w, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(w || '').trim();
}
function entwurfKey_(name, week) {
  return String(name || '').trim().toLowerCase() + '|' + entwurfWeek_(week);
}
/** Entwurf speichern (json leer/null = Entwurf löschen). Wird von der App mit Verzögerung aufgerufen. */
function saveEntwurf(nr, name, week, json) {
  if (!name || !week) return { ok: false };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = entwurfSheet_(), v = sh.getDataRange().getValues(), key = entwurfKey_(name, week), row = -1;
    for (var i = 1; i < v.length; i++) if (entwurfKey_(v[i][1], v[i][2]) === key) { row = i + 1; break; }
    if (!json) { if (row > 0) sh.deleteRow(row); return { ok: true, deleted: true }; }
    if (row < 0) row = sh.getLastRow() + 1;
    // Spalten 1-4 als Text erzwingen, sonst macht Sheets aus "2026-09-14" ein Datum und aus der Nr. eine Zahl
    sh.getRange(row, 1, 1, 4).setNumberFormat('@');
    sh.getRange(row, 1, 1, 5).setValues([[String(nr || ''), String(name), String(week), String(json), new Date()]]);
    return { ok: true, row: row };
  } finally { lock.releaseLock(); }
}
/** Alle Entwürfe eines Mitarbeiters: { "YYYY-MM-DD": { json, saved } }. */
function loadEntwuerfe(name) {
  var out = {}, key = String(name || '').trim().toLowerCase();
  if (!key) return out;
  var v = entwurfSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][1] || '').trim().toLowerCase() !== key) continue;
    var week = entwurfWeek_(v[i][2]), saved = v[i][4];
    out[week] = { json: String(v[i][3] || ''), saved: (saved instanceof Date) ? saved.toISOString() : String(saved || '') };
  }
  return out;
}
function deleteEntwurf_(name, week) { try { saveEntwurf('', name, week, null); } catch (e) {} }

/** E-Mail-Adresse eines Mitarbeiters aus dem Stammdaten-Reiter (Spalte C). */
function emailForName_(name) {
  var v = ss_().getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  var key = String(name || '').trim().toLowerCase();
  for (var i = 1; i < v.length; i++) {
    if (trim_(v[i][1]).toLowerCase() === key) {
      var mail = trim_(v[i][2]);
      return /@/.test(mail) ? mail : '';
    }
  }
  return '';
}

/**
 * Baut das Wochenrapport-PDF (analog Papierformular EIT.swiss):
 * Stundentabelle pro Auftrag, Protokoll Arbeitszeiten (Beginn/Pausen/Ende, GAV
 * Art. 20/28 bzw. ArG) und die Unterschrift des Mitarbeiters. Ablage im
 * Drive-Ordner "Wochenrapport PDFs", Rückgabe: das Drive-File.
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
    '<table style="width:80%;border-collapse:collapse"><tr style="background:#EDF3F8">' +
    td('<b>Tag</b>') + td('<b>Beginn</b>', 'text-align:center') +
    td('<b>Pausen</b>', 'text-align:center') +
    td('<b>Ende</b>', 'text-align:center') + td('<b>Stunden</b>', 'text-align:right') + '</tr>';
  (payload.protokoll || []).forEach(function (p) {
    // Alle Pausen (1–3) in einer Spalte; abwärtskompatibel zu pauseVon/pauseBis.
    var pausen = [p.pause1 || ((p.pauseVon && p.pauseBis) ? (p.pauseVon + ' – ' + p.pauseBis) : ''),
                  p.pause2 || '', p.pause3 || ''].filter(function (x) { return x; }).join(', ');
    h += '<tr>' + td(esc_(p.tag) + ' ' + esc_(p.datum)) + td(esc_(p.start), 'text-align:center') +
         td(pausen ? esc_(pausen) : '–', 'text-align:center') +
         td(esc_(p.ende), 'text-align:center') + td(num(p.stunden), 'text-align:right') + '</tr>';
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
  return folder.createFile(pdf);   // File-Objekt; URL/Blob zieht der Aufrufer
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
  var mit = readMitarbeiter_();

  var data = ss_().getSheetByName(SHEET_RAPPORTE).getDataRange().getValues();
  var lines = ['Datum;Projektnummer;Mitarbeiter Nummer;Mitarbeiter;Bemerkung;Kostenart-Nummer;Dauer'];
  var warn = {}, count = 0;

  for (var i = 1; i < data.length; i++) {
    var maNr = trim_(data[i][1]), ma = trim_(data[i][2]);
    // Sheets wandelt "11.08.2026" beim Schreiben oft in ein Date-Objekt um –
    // parseAnyDate_ akzeptiert beides; danach auf 12:00 normalisieren, damit
    // der Von/Bis-Vergleich (ebenfalls 12:00) den Randtag nicht ausschliesst.
    var datum = parseAnyDate_(data[i][5]);
    if (datum) datum = new Date(datum.getFullYear(), datum.getMonth(), datum.getDate(), 12, 0, 0);
    var auftrag = trim_(data[i][6]);
    var stunden = Number(data[i][7]);
    if (!datum || !auftrag || !(stunden > 0)) continue;
    if (datum < von || datum > bis) continue;

    var rec = mit[ma.toLowerCase()] || {};
    if (!maNr) maNr = rec.nr || '';
    if (!maNr) warn['Mitarbeiter-Nr. fehlt: ' + ma] = 1;

    var istAbsenz = kosten.absenz.hasOwnProperty(auftrag);
    var projektnummer = istAbsenz ? '' : auftrag;
    var bemerkung = istAbsenz ? auftrag : '';
    // Produktive Stunden: Kostenart je Mitarbeiter (Spalte D im Mitarbeiter-Reiter),
    // Fallback auf die Zeile "Arbeit" im Kostenarten-Reiter.
    var kostenart = istAbsenz ? kosten.absenz[auftrag] : (rec.kostenart || kosten.arbeit);
    if (!kostenart) warn['Kostenart fehlt für: ' + (istAbsenz ? auftrag : ('Arbeit bei ' + ma))] = 1;

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

/* ------------------------------------------------------------------
 * Ferienanträge und Absenzmeldungen (ersetzt die Papierformulare
 * "Ferienantrag" rosa und "Absenzmeldung" grün).
 * Ablauf: Mitarbeiter stellt Antrag -> Mail-Hinweis an GL -> GL meldet
 * sich in der App mit Name + PIN an und bewilligt/lehnt ab -> Mail an
 * den Mitarbeiter. Bewilligte Tage werden in der App-Woche vorbefüllt.
 * ------------------------------------------------------------------ */
var ANTRAG_HEADER = ['ID', 'Eingang', 'Typ', 'Mitarbeiter Nr.', 'Mitarbeiter', 'Art / Grund', 'Von', 'Bis',
                     'Tage', 'Halber Tag', 'Abzug (halber Tag)', 'Bemerkung', 'Aufgebot abgegeben',
                     'Soldmeldekarte abgegeben', 'Status', 'Entscheid am', 'Entscheid von', 'Kommentar GL',
                     'Ausbezahlt (LB)', 'Ausgebucht (Sekretariat)'];
var A = { ID: 0, EINGANG: 1, TYP: 2, NR: 3, MA: 4, ART: 5, VON: 6, BIS: 7, TAGE: 8, HALB: 9, ABZUG: 10,
          BEM: 11, AUFGEBOT: 12, SOLD: 13, STATUS: 14, ENT_AM: 15, ENT_VON: 16, KOMMENTAR: 17 };

function antragSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(SHEET_ANTRAG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ANTRAG);
    sh.getRange(1, 1, 1, ANTRAG_HEADER.length).setValues([ANTRAG_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('G:H').setNumberFormat('@');   // Datum als Text JJJJ-MM-TT, keine Umwandlung durch Sheets
  }
  return sh;
}
function antragRows_() {
  var v = antragSheet_().getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) if (v[i][A.ID]) out.push({ row: i + 1, v: v[i] });
  return out;
}
function antragObj_(r) {
  var v = r.v;
  return { id: String(v[A.ID]), eingang: fmtTs_(v[A.EINGANG]), typ: String(v[A.TYP]), ma: String(v[A.MA]),
           art: String(v[A.ART]), von: isoText_(v[A.VON]), bis: isoText_(v[A.BIS]), tage: Number(v[A.TAGE]) || 0,
           halb: String(v[A.HALB]) === 'ja', abzug: String(v[A.ABZUG] || ''), bem: String(v[A.BEM] || ''),
           aufgebot: String(v[A.AUFGEBOT]) === 'ja', sold: String(v[A.SOLD]) === 'ja',
           status: String(v[A.STATUS]), entAm: fmtTs_(v[A.ENT_AM]), entVon: String(v[A.ENT_VON] || ''),
           kommentar: String(v[A.KOMMENTAR] || '') };
}
function isoText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
function fmtTs_(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : String(v || '');
}
function isoToDate_(iso) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  return m ? new Date(+m[1], +m[2] - 1, +m[3], 12, 0, 0) : null;
}
function chFromIso_(iso) { var d = isoToDate_(iso); return d ? fmtChDate_(d) : String(iso || ''); }

/** Gesetzliche/betriebliche Feiertage Kanton Zürich (inkl. 1. Mai laut Betriebsweisung arbeitsfrei). */
function feiertage_(year) {
  var a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4), e = b % 4,
      f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30,
      i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451),
      month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1;
  var ostern = new Date(year, month - 1, day, 12);
  function off(n) { var x = new Date(ostern); x.setDate(x.getDate() + n); return isoDate_(x); }
  return [year + '-01-01', year + '-01-02', off(-2), off(1), year + '-05-01', off(39), off(50),
          year + '-08-01', year + '-12-25', year + '-12-26'];
}
/** Arbeitstage Mo–Fr ohne Feiertage zwischen zwei ISO-Daten (inklusive). */
function arbeitstage_(vonIso, bisIso) {
  var d = isoToDate_(vonIso), end = isoToDate_(bisIso), n = 0, cache = {};
  if (!d || !end || end < d) return 0;
  while (d <= end) {
    var y = d.getFullYear(); if (!cache[y]) cache[y] = feiertage_(y);
    var wd = d.getDay();
    if (wd >= 1 && wd <= 5 && cache[y].indexOf(isoDate_(d)) < 0) n++;
    d.setDate(d.getDate() + 1);
  }
  return n;
}
function glEmpfaenger_() {
  var map = readMitarbeiter_(), out = [];
  Object.keys(map).forEach(function (k) { var r = map[k];
    if (/^gl$/i.test(r.rolle || '') && /@/.test(r.email || '')) out.push(r.email); });
  return out;
}
function appUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

/**
 * Neuer Antrag. p = { ma, pin, typ: 'Ferien'|'Absenz', art, von, bis (ISO), halb, abzug: 'Ferien'|'ÜZ',
 *                     bem, aufgebot, sold }
 */
function submitAntrag(p) {
  if (!p || !p.ma) throw new Error('Bitte zuerst den Namen wählen.');
  requirePin_(p.ma, p.pin);
  var rec = readMitarbeiter_()[String(p.ma).trim().toLowerCase()] || {};
  var typ = p.typ === 'Absenz' ? 'Absenz' : 'Ferien';
  var von = String(p.von || ''), bis = String(p.bis || p.von || '');
  if (!isoToDate_(von) || !isoToDate_(bis)) throw new Error('Bitte Von- und Bis-Datum wählen.');
  if (bis < von) throw new Error('Das Bis-Datum liegt vor dem Von-Datum.');
  // Anzahl Tage trägt der Mitarbeiter selbst ein (Pensum, Teilzeit) – nur Plausibilität prüfen.
  var tage = Number(String(p.tage || '').replace(',', '.'));
  if (!(tage > 0) || Math.round(tage * 2) !== tage * 2) throw new Error('Bitte die Anzahl Arbeitstage eintragen (halbe Tage mit .5).');
  var kalTage = Math.round((isoToDate_(bis) - isoToDate_(von)) / 86400000) + 1;
  if (tage > kalTage) throw new Error('Mehr Tage eingetragen (' + tage + ') als der Zeitraum hat (' + kalTage + ').');
  var halb = tage === 0.5 && von === bis;
  var art = typ === 'Ferien' ? 'Ferien' : String(p.art || '').trim();
  if (!art) throw new Error('Bitte den Abwesenheitsgrund wählen.');

  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  var id;
  try {
    var sh = antragSheet_();
    id = 'A' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss') +
         Math.floor(Math.random() * 90 + 10);
    var row = [id, new Date(), typ, rec.nr || '', rec.name || p.ma, art, von, bis, tage, halb ? 'ja' : 'nein',
               halb ? (p.abzug === 'ÜZ' ? 'ÜZ-Konto' : 'Ferienkonto') : '', String(p.bem || '').slice(0, 300),
               typ === 'Absenz' ? (p.aufgebot ? 'ja' : 'nein') : '', typ === 'Absenz' ? (p.sold ? 'ja' : 'nein') : '',
               'offen', '', '', '', '', ''];
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 7, 1, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
  } finally { lock.releaseLock(); }

  var an = glEmpfaenger_(), mailErr = '';
  if (an.length) {
    try {
      var zeit = halb ? chFromIso_(von) + ' (halber Tag)' : chFromIso_(von) + ' – ' + chFromIso_(bis);
      MailApp.sendEmail({
        to: an.join(','),
        subject: (typ === 'Ferien' ? 'Ferienantrag' : 'Absenzmeldung') + ': ' + (rec.name || p.ma) + ', ' + zeit,
        body: 'Neuer ' + (typ === 'Ferien' ? 'Ferienantrag' : 'Absenzmeldung (' + art + ')') + ' von ' +
              (rec.name || p.ma) + ':\n\n' + zeit + ', ' + String(tage).replace('.', ',') + ' Tag(e)' +
              (p.bem ? '\nBemerkung: ' + p.bem : '') +
              '\n\nBitte in der App mit Name und PIN anmelden und unter "Freigabe" bewilligen oder ablehnen:\n' +
              appUrl_() + '\n\nDiese Nachricht wurde automatisch erstellt.'
      });
    } catch (e) { mailErr = String(e && e.message || e); console.error('Antrag-Mail an GL: ' + mailErr); }
  }
  return { ok: true, id: id, tage: tage, glMail: an.length > 0 && !mailErr, mailErr: mailErr };
}

/** Eigene Anträge (neueste zuerst, max. 30). */
function getMeineAntraege(name, pin) {
  requirePin_(name, pin);
  var key = String(name || '').trim().toLowerCase();
  return antragRows_().filter(function (r) { return String(r.v[A.MA]).trim().toLowerCase() === key; })
    .map(antragObj_).reverse().slice(0, 30);
}

/** Offenen Antrag zurückziehen (nur eigene; bewilligte nur über das Büro). */
function zurueckziehenAntrag(name, pin, id) {
  requirePin_(name, pin);
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var r = antragRows_().filter(function (x) { return String(x.v[A.ID]) === String(id); })[0];
    if (!r || String(r.v[A.MA]).trim().toLowerCase() !== String(name).trim().toLowerCase())
      throw new Error('Antrag nicht gefunden.');
    if (r.v[A.STATUS] !== 'offen') throw new Error('Nur offene Anträge können zurückgezogen werden. Bitte im Büro melden.');
    antragSheet_().getRange(r.row, A.STATUS + 1, 1, 2).setValues([['zurückgezogen', new Date()]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** GL-Ansicht: offene Anträge plus Entscheide der letzten 30 Tage, mit Überschneidungen. */
function getFreigabeListe(name, pin) {
  var chk = requirePin_(name, pin);
  if (!chk.gl) throw new Error('Keine Berechtigung (nur GL).');
  var all = antragRows_().map(antragObj_);
  var aktiv = all.filter(function (a) { return a.status === 'offen' || a.status === 'bewilligt'; });
  var grenze = new Date(); grenze.setDate(grenze.getDate() - 30);
  function overlaps(a) {
    return aktiv.filter(function (b) { return b.id !== a.id && b.ma !== a.ma && b.von <= a.bis && b.bis >= a.von; })
      .map(function (b) { return b.ma + ' (' + chFromIso_(b.von) + '–' + chFromIso_(b.bis) + ', ' + b.status + ')'; });
  }
  var offen = all.filter(function (a) { return a.status === 'offen'; })
    .map(function (a) { a.ueberschneidung = overlaps(a); return a; });
  var erledigt = antragRows_().filter(function (r) {
      var am = r.v[A.ENT_AM]; return r.v[A.STATUS] !== 'offen' && am instanceof Date && am >= grenze; })
    .map(antragObj_).reverse().slice(0, 30);
  return { offen: offen, erledigt: erledigt };
}

/** GL entscheidet: entscheid = 'bewilligt' | 'abgelehnt'. */
function entscheidAntrag(name, pin, id, entscheid, kommentar) {
  var chk = requirePin_(name, pin);
  if (!chk.gl) throw new Error('Keine Berechtigung (nur GL).');
  if (entscheid !== 'bewilligt' && entscheid !== 'abgelehnt') throw new Error('Ungültiger Entscheid.');
  if (entscheid === 'abgelehnt' && !String(kommentar || '').trim()) throw new Error('Bitte bei einer Ablehnung kurz begründen.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  var a;
  try {
    var r = antragRows_().filter(function (x) { return String(x.v[A.ID]) === String(id); })[0];
    if (!r) throw new Error('Antrag nicht gefunden.');
    if (r.v[A.STATUS] !== 'offen') throw new Error('Bereits erledigt: ' + r.v[A.STATUS] + ' von ' + (r.v[A.ENT_VON] || '–') + '.');
    var gl = (readMitarbeiter_()[String(name).trim().toLowerCase()] || {}).name || name;
    antragSheet_().getRange(r.row, A.STATUS + 1, 1, 4)
      .setValues([[entscheid, new Date(), gl, String(kommentar || '').slice(0, 300)]]);
    r.v[A.STATUS] = entscheid; r.v[A.ENT_VON] = gl; r.v[A.KOMMENTAR] = kommentar || '';
    a = antragObj_(r);
  } finally { lock.releaseLock(); }
  try {
    var mail = emailForName_(a.ma);
    if (mail) {
      var zeit = a.halb ? chFromIso_(a.von) + ' (halber Tag)' : chFromIso_(a.von) + ' – ' + chFromIso_(a.bis);
      MailApp.sendEmail({
        to: mail,
        subject: (a.typ === 'Ferien' ? 'Ferienantrag ' : 'Absenzmeldung ') + entscheid + ': ' + zeit,
        body: 'Hallo ' + a.ma + '\n\nDein ' + (a.typ === 'Ferien' ? 'Ferienantrag' : 'Absenzmeldung (' + a.art + ')') +
              ' für ' + zeit + ' wurde von ' + a.entVon + ' ' + entscheid + '.' +
              (a.kommentar ? '\nKommentar: ' + a.kommentar : '') +
              (entscheid === 'bewilligt' ? '\n\nDie Tage erscheinen in der App automatisch in deiner Wochenerfassung.' : '') +
              '\n\nDiese Nachricht wurde automatisch erstellt.'
      });
    }
  } catch (e) {}
  return { ok: true };
}

/**
 * Bewilligte Abwesenheiten einer Woche für die Vorbefüllung der App:
 * [{ auftrag: 'Ferien Bezug', h: [Mo..So] }, ...]
 */
function getBewilligteWoche(name, pin, weekIso) {
  requirePin_(name, pin);
  var mon = isoToDate_(weekIso); if (!mon) return [];
  var key = String(name || '').trim().toLowerCase(), labels = Object.keys(readKostenarten_().absenz);
  function label(a) {
    if (a.typ === 'Ferien') {
      if (a.halb && a.abzug === 'ÜZ-Konto') return labels.filter(function (l) { return /kompensation/i.test(l); })[0] || 'Kompensation Überstunden (frei)';
      return labels.filter(function (l) { return /^ferien bezug$/i.test(l); })[0] || 'Ferien Bezug';
    }
    return a.art;
  }
  var byLabel = {}, feiern = feiertage_(mon.getFullYear()).concat(feiertage_(mon.getFullYear() + 1));
  antragRows_().forEach(function (r) {
    if (r.v[A.STATUS] !== 'bewilligt' || String(r.v[A.MA]).trim().toLowerCase() !== key) return;
    var a = antragObj_(r), lb = label(a);
    for (var i = 0; i < 5; i++) {
      var d = new Date(mon); d.setDate(d.getDate() + i);
      var iso = isoDate_(d);
      if (iso < a.von || iso > a.bis || feiern.indexOf(iso) >= 0) continue;
      if (!byLabel[lb]) byLabel[lb] = [0, 0, 0, 0, 0, 0, 0];
      byLabel[lb][i] = a.halb ? SOLL_STUNDEN / 10 : SOLL_STUNDEN / 5;
    }
  });
  return Object.keys(byLabel).map(function (k) { return { auftrag: k, h: byLabel[k] }; });
}

/** Täglicher Hinweis an die GL, wenn Anträge seit mehr als 3 Tagen offen sind. */
function erinnerungAntraege() {
  var grenze = new Date(); grenze.setDate(grenze.getDate() - 3);
  var alt = antragRows_().filter(function (r) {
    return r.v[A.STATUS] === 'offen' && r.v[A.EINGANG] instanceof Date && r.v[A.EINGANG] < grenze; }).map(antragObj_);
  var an = glEmpfaenger_();
  if (!alt.length || !an.length) return;
  MailApp.sendEmail({
    to: an.join(','),
    subject: alt.length + ' offene Anträge warten auf Freigabe',
    body: alt.map(function (a) { return '- ' + a.ma + ': ' + a.typ + ' ' + chFromIso_(a.von) + ' – ' + chFromIso_(a.bis) +
                                  ' (eingegangen ' + a.eingang + ')'; }).join('\n') +
          '\n\nFreigabe in der App unter "Freigabe": ' + appUrl_()
  });
}
/** Menü: zeigt, wer die Antrag-Mails bekommt, und schickt eine Testmail dorthin. */
function testAntragMail() {
  var ui = SpreadsheetApp.getUi(), an = glEmpfaenger_();
  if (!an.length) { ui.alert('Niemand gefunden: Im Reiter Mitarbeiter braucht es in Spalte F "GL" und in Spalte C eine E-Mail.'); return; }
  try {
    MailApp.sendEmail({ to: an.join(','), subject: 'Test Ferienantrag-Mail',
      body: 'Testmail aus dem Wochenrapport. Diese Adressen erhalten neue Ferienanträge und Absenzmeldungen:\n' + an.join('\n') });
    ui.alert('Testmail gesendet an:\n' + an.join('\n') + '\n\nVerbleibendes Mail-Kontingent heute: ' + MailApp.getRemainingDailyQuota());
  } catch (e) { ui.alert('Mail fehlgeschlagen: ' + (e && e.message || e)); }
}

function setupErinnerung() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'erinnerungAntraege') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('erinnerungAntraege').timeBased().everyDays(1).atHour(7).create();
  try { SpreadsheetApp.getUi().alert('Tägliche Erinnerung (ca. 07:00) für offene Anträge ist eingerichtet.'); } catch (e) {}
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

/** name (klein) -> { nr, name, email, kostenart, pin, rolle } aus dem Mitarbeiter-Reiter (Spalten A–F). */
function readMitarbeiter_() {
  var map = {};
  var v = ss_().getSheetByName(SHEET_MITARBEITER).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var nr = trim_(v[i][0]), name = trim_(v[i][1]), kostenart = trim_(v[i][3]), pin = trim_(v[i][4]);
    if (name) map[name.toLowerCase()] = { nr: nr, name: name, email: trim_(v[i][2]), kostenart: kostenart,
                                          pin: pin, rolle: trim_(v[i][5]) };
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
