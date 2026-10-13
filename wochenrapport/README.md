# Digitaler Wochenrapport

Ersetzt das handschriftliche Wochenrapport-Formular (EIT.swiss 631) durch eine
zentrale, digitale Erfassung. Ziel: Die ~10 Stunden manuelles Abtippen in das
Zeiterfassungssystem (Messerli) entfallen.

## Wie es funktioniert

- **40 Mitarbeiter** öffnen **einen Link** (ein Google-Web-Formular).
- Jeder wählt seinen Namen, die Woche und erfasst pro Zeile eine **Auftragsnummer**
  (mit Sofort-Vorschlägen beim Tippen) und die Stunden pro Tag.
- Mit **„Woche absenden"** landen alle Eingaben **zentral** in einem Google Sheet.
- Die Administration exportiert aus dem Sheet **eine CSV-Datei** für den Messerli-Import.

Alles läuft parallel; ein Lock verhindert Konflikte beim gleichzeitigen Absenden.
Vor dem Absenden werden Eingaben lokal als Entwurf zwischengespeichert, damit nichts
verloren geht.

```
Mitarbeiter  ──►  Web-Formular  ──►  Google Sheet "Rapporte"  ──►  CSV  ──►  Messerli
(40 parallel)     (ein Link)         (zentrale Ablage)              (1 Import)
```

## Warum Google Sheet + Apps Script

- **Kostenlos** und läuft im vorhandenen Google-Konto.
- Google **hostet die Seite** (kein eigener Server nötig).
- Die zentrale Ablage ist eine **ganz normale Tabelle** – Auswertung und Export mit Bordmitteln.
- Projekte und Namen pflegst du per **Copy-Paste aus Messerli** direkt im Sheet.

## Einrichtung (einmalig, ~15 Minuten)

1. **Google Sheet anlegen:** [sheets.new](https://sheets.new) öffnen, Datei z. B. „Wochenrapport" nennen.
2. **Apps Script öffnen:** im Sheet Menü **Erweiterungen → Apps Script**.
3. **Backend einfügen:** den vorhandenen Inhalt löschen und `apps-script/Code.gs` aus diesem
   Repo komplett hineinkopieren.
4. **Formular als HTML-Datei anlegen:** links auf **+ → HTML**, die Datei exakt `Index` nennen
   (Apps Script ergänzt `.html` automatisch), Inhalt aus `apps-script/Index.html` hineinkopieren.
5. **Speichern** (Disketten-Symbol).
6. **Reiter erzeugen:** Funktion `ensureSheets_` einmal ausführen (oben Funktion wählen → **Ausführen**),
   beim ersten Mal die Berechtigungen bestätigen. Es entstehen die Reiter `Rapporte`, `Projekte`, `Mitarbeiter`.
7. **Veröffentlichen:** **Bereitstellen → Neue Bereitstellung → Typ: Web-App**.
   - *Ausführen als:* **Ich**
   - *Zugriff:* **Jeder** (bzw. „Jeder mit Google-Konto", je nach gewünschter Kontrolle)
   - **Bereitstellen** klicken, den Zugriff autorisieren, die **Web-App-URL** kopieren.
8. **URL an die Mitarbeiter** verteilen. Fertig.

> Nach Änderungen am Code: **Bereitstellen → Bereitstellungen verwalten → Bearbeiten →
> Version: Neu → Bereitstellen** (sonst bleibt die alte Version aktiv).

## Stammdaten pflegen (laufend)

Im Google Sheet:

- **Reiter `Projekte`:** Spalte A = **Projektnummer exakt wie in MesserliPROJEKT**
  (z. B. `26-26581.001`), Spalte B = Bezeichnung (optional, hilft nur bei der Auswahl).
  Einfach aus Messerli kopieren und einfügen.
- **Reiter `Mitarbeiter`:** Spalte A = **Mitarbeiter-Nummer exakt wie in Messerli**,
  Spalte B = Name. Die Datei `mitarbeiter-liste.tsv` in diesem Repo enthält alle 69
  Mitarbeiter mit Nummern zum direkten Einfügen (Stand JahresRapport 20.07.2026).
- **Reiter `Kostenarten`:** Spalte A = Typ (`Arbeit` sowie die Absenzgründe),
  Spalte B = **Kostenart-Nummer exakt wie in Messerli**. `Arbeit` gilt für alle
  produktiven Projektstunden, die Absenzgründe haben je eigene Nummern.
  `Kompensation Überstunden (frei)` mit Kostenart `K-ÜS` wird beim nächsten Öffnen
  der App automatisch ergänzt, falls die Zeile fehlt.

- **Spalte F `Rolle` im Reiter `Mitarbeiter`:** `GL` bei allen, die Ferienanträge und
  Absenzmeldungen bewilligen dürfen. Eine GL-Person braucht zwingend einen PIN (Spalte E)
  und eine E-Mail-Adresse (Spalte C). PINs nie in dieses Repo schreiben.

Änderungen sind sofort wirksam – kein erneutes Veröffentlichen nötig.

## Ferienantrag und Absenzmeldung (ab Version 5)

Ersetzt die Papierformulare. Ablauf:

1. Mitarbeiter meldet sich in der App an und öffnet den Reiter **Ferien / Absenz**.
   Ferienantrag oder Absenzmeldung (Militär/Zivilschutz, Weiterbildung, Schule, Kompensation …),
   Von/Bis, **Anzahl Arbeitstage selbst eintragen** (wegen Pensum/Teilzeit, halbe Tage mit .5;
   bei 0.5 Tag Ferien: Abzug Ferien- oder ÜZ-Konto wählen), Bemerkung, bei Absenzen
   Aufgebot/Soldmeldekarte abgegeben.
2. Der Antrag landet im Reiter **`Anträge`** (Status `offen`), alle GL-Personen erhalten eine
   Mail mit dem Link zur App.
3. Die GL meldet sich mit Namen + PIN an und sieht den Reiter **Freigabe**: offene Anträge mit
   Hinweis, wer im gleichen Zeitraum schon abwesend ist. *Bewilligen* oder *Ablehnen*
   (Ablehnung nur mit Begründung). Status, Datum und Name der GL werden gespeichert,
   der Mitarbeiter erhält eine Mail.
4. Solange ein Antrag offen ist, kann ihn der Mitarbeiter zurückziehen.
5. Bewilligte Tage erscheinen automatisch im Wochenrapport, wenn die Woche noch leer ist
   (Ferien Bezug 8 h, halber Tag 4 h, halber Tag ab ÜZ-Konto als Kompensation).
6. Lohnbuchhaltung/Sekretariat haken im Sheet die Spalten *Ausbezahlt (LB)* und
   *Ausgebucht (Sekretariat)* ab.

Sicherheit: GL-Rechte prüft nur der Server nach korrektem PIN. Nach 5 falschen PIN-Eingaben
ist der Name 15 Minuten gesperrt. Die Rolle wird nie an den Browser geschickt.

Erinnerung: Menü **Wochenrapport → Erinnerung offene Anträge einrichten (täglich)** einmal
ausführen. Danach erhält die GL jeden Morgen um 7 Uhr eine Mail, wenn Anträge länger als
3 Tage offen sind.

## Kontrolle: Wer hat eingereicht?

Menü **Wochenrapport → Kontrolle aktualisieren** baut den Reiter `Kontrolle`:
alle Mitarbeiter untereinander, die Kalenderwochen (KW) nebeneinander, in den
Zellen die eingereichten Stunden.

- 🟩 Grün = Woche vollständig (≥ Sollzeit 40 h)
- 🟨 Gelb = eingereicht, aber unter 40 h (prüfen)
- 🟥 Rot = nichts eingereicht

Damit siehst du auf einen Blick, wem du nachlaufen musst, bevor du den
Messerli-Export machst.

## Export nach Messerli

Das Importformat wurde mit dem Messerli-Support geklärt (Ticket BUS-SUP-27446,
`Muster_Arbeitszeitimport-Vorlage.csv`):

```
Datum;Projektnummer;Mitarbeiter Nummer;Mitarbeiter;Bemerkung;Kostenart-Nummer;Dauer
16.01.2026;26-26581.001;303;Markus Müller;Lampe montieren;102;08:00
```

- Spaltenreihenfolge darf abweichen; ohne Von/Bis-Zeiten genügt die Spalte **Dauer**.
- Projektnummern, Mitarbeiternummern und Kostenartnummern müssen **exakt** den
  Nummern aus MesserliPROJEKT entsprechen.

Im Sheet: Menü **Wochenrapport → Messerli-Export erstellen …**, Zeitraum eingeben
(von/bis als TT.MM.JJJJ) – die CSV wird im Drive abgelegt und der Link angezeigt.
Dabei gilt:

- `Dauer` wird aus den Dezimalstunden umgerechnet (8.5 → `08:30`).
- Produktive Stunden: `Projektnummer` = Auftrag-Nr., Kostenart = Zeile `Arbeit`.
- Absenzen (Ferien, Krank, …): `Projektnummer` leer, Absenzgrund in `Bemerkung`,
  Kostenart aus dem Reiter `Kostenarten`.
- Fehlende Mitarbeiter- oder Kostenart-Nummern werden beim Export als Warnung gemeldet.

Der Reiter **`Rapporte`** (Rohdaten) hat die Spalten:

| Zeitstempel | Mitarbeiter Nr. | Mitarbeiter | Woche vom | Rapport Nr. | Datum | Auftrag Nr. | Stunden |
|---|---|---|---|---|---|---|---|

Laut Messerli-Support dauert das Einrichten und Testen des Imports auf deren Seite
ca. 2–3 h (Fr. 180.–/h); weitere Module werden nicht benötigt.

## Struktur dieses Repos

```
wochenrapport/
├─ apps-script/
│  ├─ Code.gs             Backend: Formular, zentrale Ablage, Messerli-CSV-Export
│  └─ Index.html          Web-Formular (Stundenerfassung, Ferien/Absenz-Antrag, GL-Freigabe)
├─ mitarbeiter-liste.tsv  69 Mitarbeiter (Nr. + Name) zum Einfügen in den Reiter "Mitarbeiter"
└─ README.md              diese Anleitung
```

## Offene Punkte / nächste Ausbaustufe

- **Kostenart-Nummern** im Reiter `Kostenarten` eintragen (aus MesserliPROJEKT übernehmen).
- **Projektnummern** im Reiter `Projekte` einfügen (exakt wie in Messerli, z. B. `26-26581.001`).
- Import auf Messerli-Seite einrichten/testen lassen (Angebot Support: ca. 2–3 h à Fr. 180.–).
- Optional: Visum-Schritt für Wochenrapporte durch Vorgesetzte.
- Geplant: aktueller Stunden- und Feriensaldo in der App (aus Messerli-Excel).
