# Analyse des Repositories „Flow Framework 2“

## 1) Gesamtbild
- Das Repository enthält den Export einer Excel-VBA-Arbeitsmappe (`FlowFramework2.xlsb`) plus alle exportierten VBA-Module/Klassen, Metadaten-Dateien und Dokumentation.
- Architekturprinzip laut Doku: Trennung zwischen **Framework-Core (`f_`)**, **app-spezifischem Framework-Bereich (`af_`)** und **App-Code (`a_`)**.
- Zusätzlich existiert **DEV-Core (`DEV_`)** für Entwicklung/Tests sowie zwei alte Snapshot-Ordner als **DEPRECATED**.

## 2) Datei-Inventar (vollständig abgedeckt)
Analysiert wurden:
- Root: alle `.bas`, `.cls`, `.md`, `.json`, sowie `FlowFramework2.xlsb` (nur als Binär-Artefakt identifiziert).
- Unterordner: `independent-features-DEPRECATED/*` und `ff2s-little-sis-DEPRECATED/*` vollständig als Legacy/Referenzbestand gesichtet.

## 3) Funktionsweise des Frameworks

### 3.1 Laufzeit-Grundgerüst
1. Entry-Level-Prozeduren rufen am Anfang `f_p_StartProcessing` auf.
2. Dadurch werden globale Objekte initialisiert (`oC_f_p_FrameworkSettings`, Fehler-Collection, optional DEV-Init).
3. Je nach Processing-Mode werden z. B. ScreenUpdating/Calculation geschaltet oder app-spezifische Hooks ausgeführt.
4. Am Ende wird `f_p_EndProcessing` aufgerufen und der Modus sauber zurückgesetzt.

### 3.2 Fehlerbehandlungsmuster
- Fast alle „echten“ Funktionen folgen einem Template mit:
  - `oC_Me As f_C_CallParams` für Prozedur-Kontext,
  - `Try/Catch/Finally`-ähnlichem VBA-Muster (`On Error GoTo Catch`),
  - zentralen Framework-Funktionen `f_p_RegisterError` + `f_p_HandleError`,
  - optionalen app-spezifischen Error-Hooks (`af_p_Hook_ErrorHandling_*`).
- `f_C_Error` und `f_C_CallParams` kapseln Fehlerdaten + Call-Kontext.

### 3.3 Modi (Debug / Dev / Maintenance)
- Die Modi werden über benannte Zellen in `af_wks_Settings` persistiert.
- `af_C_AppModes` steuert Sichtbarkeit von Worksheets je Modus:
  - Development-Sheets werden sichtbar/very hidden,
  - Maintenance-Sheets ebenfalls separat gesteuert.
- Passwortprüfung für Dev/Maintenance ist vorgesehen (derzeit Platzhalter-Passwörter).

### 3.4 Konfiguration über Settings-Sheets + Names
- `f_C_Settings` liest Kernwerte (Version, Flags) über Excel-Namen.
- `f_C_SettingsSheet` + `f_C_Setting` bilden Settings-Zeilen aus Worksheets als Objekte ab.
- Exportdateien (`SettingsSheet-*.json`, `Names.json`, `WorksheetNames.json`, `References.json`) dienen der Versionskontrolle und Rekonstruktion relevanter Workbook-Metadaten.

### 3.5 Deployment / DEV-Entfernung
- `f_C_Deploy.bSaveAsProdAndRemoveDEVModules` erstellt eine Produktivkopie und entfernt DEV-Komponenten per VBIDE-Automation.
- `f_pM_EntryLevel.f_p_DeployWorkbook` kapselt den Deploy-Use-Case als Entry-Level-Sub.

### 3.6 Version-Control-Export
- `DEV_f_C_VersionControlExport` exportiert:
  - definierte Range-Inhalte,
  - Settings-Sheets,
  - Referenzen,
  - Worksheet-Namen,
  - Names.
- `DEV_f_C_VersionControlRanges` liest aus `a_wks_VersionControlRanges`, welche Ranges exportiert werden sollen.

### 3.7 Utilities & Datenabstraktion
- `f_C_Wks`: Worksheet-Wrapper (DataRange, Header-Dictionary, CurrentRegion-Logik, Sanitizing).
- `f_C_DataRecord` + `f_I_DataRecord`: generische Datensatz-Abstraktion via Dictionary.
- Utility-Module für Namen-/Worksheet-Zugriffe, Range-Vergleiche und Dateisystemzugriffe.
- `f_C_DriveMapper` ist für Pfad-/Netzlaufwerkszenarien vorgesehen.

### 3.8 Template Rendering (2 Renderer)
- **Renderer Lite** (`f_pM_TemplRendererLite`): ein Block + eine Repeater-Zeile (`rep_Items`), Platzhalterersetzung `{{...}}`.
- **Renderer Blocks/Lanes** (`f_pM_TemplRenderer`):
  - Blöcke (`blk_*`) mit Lanes `fix_`, `rep_`, `rel_`.
  - Dynamisches Nach-unten-Schieben von Bereichen bei Expansion.
  - Kontext-Objekt mit `header`, `repeaters`, `totals`.
- Styling (`f_pM_TemplRenderer_Styles`) wird aus `_meta`/Styles-Range gelesen; Style-Tokens kommen aus Zellkommentaren (`style:Token`) inkl. Border-Regeln.

### 3.9 UI-Einstiegspunkte
- Framework-UI (`f_M_UserInterface`) bietet Klick-Einstiegspunkte (u. a. Maintenance toggle).
- App-UI (`a_M_UserInterface`) ist bewusst leer als Erweiterungspunkt.

## 4) Was bereits sichtbar nur Skelett/Platzhalter ist
- `a_*`-Module sind größtenteils Gerüste ohne Business-Logik.
- `af_pM_ErrorHandling` enthält Hooks ohne Implementierung.
- `DEV_f_pM_Testing.DEV_f_m_RunUnitTests` ist TODO (kein echter Runner).
- `f_C_RangeArrayProcessor` enthält nur `SanitizeLeadingZeroItems`; Rest TODO.
- `f_pM_TemplatesCore` / `f_pM_TemplatesCoreCompact` sind bewusst Vorlagenmodule.
- In den Renderer-Modulen sind mehrere TODOs zur Framework-Integration/Refaktorierung.

## 5) Auffälligkeiten aus den Exportdaten
- Mehrere `SettingsSheet-*.json` sind formal ungültig (z. B. trailing commas, fehlendes Komma zwischen Objekten). Für rein menschliche Diff-Zwecke ist das tolerierbar, für maschinelles JSON-Parsing nicht.
- `Names.json`/`WorksheetNames.json` zeigen, dass die Template-Demo-Blätter und Named-Ranges für beide Renderer bereits im Workbook angelegt sind.

## 6) Was ich ohne weitere Angaben von dir nicht sicher beantworten kann
1. **Produktive Ziel-App-Logik**: Die `a_*`-Schicht ist absichtlich leer – fachliche Prozesse fehlen.
2. **Modus-Passwörter/Policies**: In `af_C_AppModes` sind Platzhalterwerte gesetzt.
3. **Gewünschtes Error-Handling-Verhalten** in den app-spezifischen Hooks (`af_pM_ErrorHandling`).
4. **Ob die JSON-Exporte absichtlich „nicht strikt JSON“** sind oder als Bug zu behandeln.
5. **Priorität der beiden Renderer** für deine konkreten Use-Cases (Lite vs. Blocks/Lanes).
6. **Geplanter Testansatz**: Der DEV-Test-Runner ist unvollständig; unklar, ob du Unit-Tests im Framework finalisieren willst oder extern testest.
7. **Verwendung der DEPRECATED-Ordner**: Snapshot-Referenz vs. noch operative Migrationsquelle.

## 7) Kurzfazit
Das Framework liefert ein sauberes VBA-Architektur-Skelett mit starkem Fokus auf:
- standardisierte Entry-/Lower-Level-Fehlerbehandlung,
- Modus- und Sichtbarkeitssteuerung,
- versionskontrollierbare Workbook-Metadaten,
- deploybare Trennung von DEV- und PROD-Bestand,
- sowie Template-basierte Dokument-/Sheet-Generierung (inkl. Styles).

Die eigentliche Fachanwendung ist erwartungsgemäß noch nicht implementiert; mehrere Teile sind explizit als Vorlage/TODO markiert.
