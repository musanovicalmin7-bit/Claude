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

- **Reiter `Projekte`:** Spalte A = Auftragsnummer, Spalte B = Bezeichnung (optional).
  Einfach aus Messerli kopieren und einfügen. Die Bezeichnung hilft nur bei der Auswahl.
- **Reiter `Mitarbeiter`:** Spalte A = Name (ein Name pro Zeile).

Änderungen sind sofort im Formular wirksam – kein erneutes Veröffentlichen nötig.

## Export nach Messerli

Der Reiter **`Rapporte`** enthält alle Buchungen mit den Spalten:

| Zeitstempel | Mitarbeiter | Woche vom | Rapport Nr. | Datum | Auftrag Nr. | Stunden |
|---|---|---|---|---|---|---|

Zum Import: gewünschten Zeitraum filtern und **Datei → Herunterladen → CSV**. Das
Spaltenformat lässt sich exakt an die Messerli-Importvorlage anpassen, sobald diese bekannt ist.

## Struktur dieses Repos

```
wochenrapport/
├─ apps-script/
│  ├─ Code.gs       Backend: liefert das Formular, sammelt Rapporte, liest Stammdaten
│  └─ Index.html    Web-Formular (Namensauswahl, Auftrag-Autocomplete, Stundenerfassung)
└─ README.md        diese Anleitung
```

## Offene Punkte / nächste Ausbaustufe

- **Messerli-Importformat:** exakte Spaltenreihenfolge/Trennzeichen der Messerli-Importvorlage
  abklären und CSV-Export daran anpassen.
- Optional: Freigabe-/Visum-Schritt durch Vorgesetzte, Wochenübersicht pro Mitarbeiter,
  automatischer CSV-Export per Menüpunkt.
- Alternative geprüft: Messerli-eigene Web-Erfassung (MesserliZERF / Web-Rapport) – falls
  lizenziert, entfällt der Import ganz.
