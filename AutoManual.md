# AutoManual – Flow Framework 2

## 1. Überblick
Flow Framework 2 ist ein Excel-VBA-Framework, das wiederverwendbare Architekturen, robuste Fehlerbehandlung und automatisierbare Deployment- sowie Versionskontroll-Prozesse bereitstellt. Kernziel ist es, Anwendungen innerhalb eines einzigen Arbeitsbuchs (**FlowFramework2.xlsb**) sauber strukturiert zu entwickeln, ohne Code direkt in Arbeitsblatt- oder Arbeitsmappen-Objekten zu hinterlegen.

Dieser Auto-Manual ergänzt das bestehende `developer-manual.md` und liefert eine systematische Referenz der Komponenten im Projektwurzelverzeichnis. Er erläutert Funktionsumfang, Erweiterungspunkte und aktuelle Baustellen der Codebasis.

## 2. Struktur und Namenskonventionen
Die Dateien im Projektwurzelverzeichnis folgen klaren Präfixen:

- `f_…`: Framework-Kern, nicht verändern (z. B. `f_pM_GlobalsCore.bas`).
- `af_…`: anpassbarer Framework-Teil mit Anwendungskontext (z. B. `af_C_AppModes.cls`).
- `a_…`: Platzhalter für Applikationscode (z. B. `a_M_UserInterface.bas`).
- `DEV_…`: Entwicklungs-Hilfen; werden für Deployments üblicherweise entfernt (z. B. `DEV_f_C_VersionControlExport.cls`).

Die Namenskonventionen für Variablen, Komponenten und Excel-Namen sind im bestehenden Developer Manual ausführlich beschrieben und bleiben unverändert gültig. Das Framework setzt stark auf ungarische Notation und Präfixe zur Erhöhung der Lesbarkeit.

## 3. Voraussetzungen und Referenzen
Für die erfolgreiche Ausführung werden die in `References.json` aufgeführten Bibliotheken benötigt, u. a. Microsoft Excel-, Office- und VBA-Standardbibliotheken sowie **Microsoft Scripting Runtime** und **Windows Script Host Object Model**, die z. B. vom Klassenmodul `f_C_DriveMapper.cls` genutzt werden.【F:References.json†L2-L41】【F:f_C_DriveMapper.cls†L1-L86】

## 4. Betriebsmodi und zentrale Einstellungen
### 4.1 Globale Framework-Einstellungen
- `f_C_Settings.cls` liest Versionsinformationen und Statusflags (Debug-, Entwicklungs- und Wartungsmodus) aus den definierten Excel-Namen in `f_wks_Settings` und `af_wks_Settings` und synchronisiert Änderungen zurück in die Tabellenblätter.【F:f_C_Settings.cls†L1-L134】
- Die Einträge werden in den JSON-Dateien `SettingsSheet-f_wks_Settings.json` und `SettingsSheet-af_wks_Settings.json` gespiegelt, was einen Blick auf Struktur und Feldnamen erlaubt.【F:SettingsSheet-f_wks_Settings.json†L1-L13】【F:SettingsSheet-af_wks_Settings.json†L1-L19】

### 4.2 Moduswechsel
- `f_pM_EntryLevel.bas` stellt öffentlich erreichbare Prozeduren bereit, um Wartungs- und Entwicklungsmodus mit Passwortabfrage zu toggeln. Die Logik nutzt `af_C_AppModes.cls`, aktiviert/deaktiviert Arbeitsblätter, setzt Namenssichtbarkeit und aktualisiert Framework-Flags.【F:f_pM_EntryLevel.bas†L1-L138】【F:af_C_AppModes.cls†L1-L206】
- `af_C_AppModes.cls` definiert konfigurierbare Passwortkonstanten sowie Sammlungen sichtbar zu schaltender Arbeitsblätter. Entwickler müssen hier ihre eigenen Entwicklungs- bzw. Wartungsblätter ergänzen.【F:af_C_AppModes.cls†L19-L116】
- Die aktuellen Moduswerte werden in `af_wks_Settings` gepflegt und sind ebenfalls im JSON-Export ersichtlich.【F:SettingsSheet-af_wks_Settings.json†L1-L19】

### 4.3 Start- und Endlogik
- `f_pM_GlobalsCore.bas` initialisiert globale Objekte (`oC_f_p_FrameworkSettings`, Fehler- und Test-Sammlungen) und kapselt Start-/Endverhalten über definierte Verarbeitungsmodi (z. B. Deaktivierung von Bildschirmaktualisierung). Es ruft optional Entwicklungs-Hooks wie `DEV_f_p_InitGlobals` auf.【F:f_pM_GlobalsCore.bas†L1-L135】
- App-spezifische Erweiterungen erfolgen in `af_pM_Globals.bas`, das eine Enum `e_af_p_ProcessingModes` sowie entsprechende Start-/Endverarbeitungsschalter bereitstellt.【F:af_pM_Globals.bas†L1-L88】

## 5. Fehlerbehandlung und Logging
- `f_pM_ErrorHandling.bas` verwaltet zentral die Fehlerregistrierung, legt Fehlerobjekte (`f_C_Error.cls`) in einer globalen Sammlung ab und schreibt ausführliche Log-Einträge (inkl. Argumentlisten) in das Arbeitsblatt `af_wks_ErrorLog`. Die Routine `s_f_p_HandledErrorDescription` erzeugt benutzerfreundliche Meldungen.【F:f_pM_ErrorHandling.bas†L1-L87】
- `af_pM_ErrorHandling.bas` liefert Erweiterungspunkte: Entwickler können eigene Fehlerenumerationen, beschreibende Texte sowie Hooks für Entry-Level- und Lower-Level-Fehler ergänzen (derzeit mit Platzhaltern markiert).【F:af_pM_ErrorHandling.bas†L1-L80】
- Die Fehlerlog-Struktur (Spalten A–I) ist in `f_pM_ErrorHandling.bas` hardcodiert und sollte bei strukturellen Änderungen am Arbeitsblatt angepasst werden.【F:f_pM_ErrorHandling.bas†L60-L87】

## 6. Templates und Entwicklungsleitlinien
- `f_pM_TemplatesCore.bas` stellt ausführlich kommentierte Schablonen für Entry-Level-Prozeduren und nicht-triviale Lower-Level-Funktionen bereit. Die Templates beinhalten bereits Error-Handling, Test-Hooks und Hinweise auf die benötigten Aufräumarbeiten. TODO-Platzhalter markieren fehlende Beispielimplementierungen, insbesondere für `f_p_TemplateSubEntryLevel` und `b_f_p_TemplateLowerLevel`.【F:f_pM_TemplatesCore.bas†L1-L142】
- `a_M_UserInterface.bas` und `a_pM_EntryLevel.bas` sind leere Container für applikationsspezifische Einstiegspunkte. Sie sollten ausschließlich Wrapper enthalten, die zu den Framework-Templates passen.【F:a_M_UserInterface.bas†L1-L24】【F:a_pM_EntryLevel.bas†L1-L27】
- `a_pM_OnChangeSubsFor_f_C_Wks.bas` dient als Sammelstelle für Ereignis-Handler, die von `f_C_Wks`-Instanzen bei Arbeitsblattänderungen aufgerufen werden können.【F:a_pM_OnChangeSubsFor_f_C_Wks.bas†L1-L11】

## 7. Einstellungen und Datenhaltung
- `f_C_SettingsSheet.cls` kapselt das Auslesen strukturierter „Settings“-Blätter: `bConstruct` parametrisiert Zeilen-/Spaltenlayout, `bGetSettingsFromSettingsSheet` baut `f_C_Setting`-Objekte und schreibt sie in bereitgestellte Collections. Entwickler müssen in den markierten Abschnitten eigenen Code ergänzen, wenn zusätzliche Verarbeitung nötig ist.【F:f_C_SettingsSheet.cls†L1-L198】
- `af_pM_Globals.bas` liefert dazu die vorkonfigurierte Sammlung relevanter Settings-Blätter (`f_wks_Settings`, `af_wks_Settings`, `a_wks_Settings`, `a_wks_VersionControlRanges`), die bereits in den JSON-Exports gespiegelt werden.【F:af_pM_Globals.bas†L47-L88】
- Für anwendungsspezifische Konfiguration steht `a_C_AppSettings.cls` bereit, das Versionsnummer und -datum aus Workbook-Namen übernimmt.【F:a_C_AppSettings.cls†L1-L23】

## 8. Datenobjekte und Utility-Klassen
### 8.1 Datenhaltung
- `f_I_DataRecord.cls` und `f_C_DataRecord.cls` definieren ein Dictionary-basiertes Datenobjekt mit Primärschlüssel. Getter/Setter geben Boolean-Rückgaben für Erfolgsmeldungen aus und unterstützen polymorphe Erweiterungen.【F:f_I_DataRecord.cls†L1-L18】【F:f_C_DataRecord.cls†L1-L64】

### 8.2 Worksheet-Kapselung
- `f_C_Wks.cls` erweitert `Worksheet` um Funktionen zur Bereichsverwaltung, Header-Dictionary-Erstellung, Ereignisweiterleitung und Sanitizing von UsedRanges. Die Backlog-Notiz erwähnt fehlende Ereignisimplementierungen für `oWks_m_Wks`. Entwickler müssen außerdem `s_prop_rw_NameOfSubToRunOnWksChange` setzen, damit die Change-Events Routing auslösen.【F:f_C_Wks.cls†L1-L205】【F:f_C_Wks.cls†L206-L292】

### 8.3 Range-Verarbeitung
- `f_C_RangeArrayProcessor.cls` enthält aktuell nur die Funktion `SanitizeLeadingZeroItems` sowie mehrere TODOs: Konstruktor, Dictionaries für Zeilen-/Spaltenindizes und optionale Namensunterstützung sind noch zu implementieren.【F:f_C_RangeArrayProcessor.cls†L1-L46】

### 8.4 Utility-Module
- `f_pM_Utilities.bas` stellt generische Helfer wie `s_f_p_MyProcedureName`, Name-/Range-Getter sowie Schlüssel-Sanitizer bereit.【F:f_pM_Utilities.bas†L1-L112】
- `f_pM_UtilitiesDev.bas` bietet Funktionen zum (De-)Aktivieren technischer Namen/Blätter und ein Debugging-Hilfslog (`f_p_PrintCallParams`).【F:f_pM_UtilitiesDev.bas†L1-L82】
- `f_pM_UtilitiesRanges.bas` vergleicht Range-Größen und Inhalte und liefert Statusmeldungen. Das Modul ist neu (Version 1.16.0).【F:f_pM_UtilitiesRanges.bas†L1-L70】
- `f_pM_UtilitiesFileSystem.bas` kapselt das Öffnen externer Arbeitsmappen sowie das Auflösen von Arbeitsblatt-CodeNamen in Objekte, inklusive Error-Handling gemäß Framework-Konventionen.【F:f_pM_UtilitiesFileSystem.bas†L1-L86】
- `f_C_DriveMapper.cls` mappt SharePoint-/Netzwerkpfade auf lokale Laufwerksbuchstaben und sorgt für automatische Unmounts beim Klassenende.【F:f_C_DriveMapper.cls†L1-L86】

## 9. Deployment und Versionskontrolle
- `f_C_Deploy.cls` speichert eine Produktionskopie (`PROD-…`) des aktuellen Arbeitsbuchs und entfernt alle Komponenten mit Präfix `DEV`. Anpassungen sind hauptsächlich in den markierten Bereichen vorzunehmen (z. B. zusätzliche Bereinigungsschritte).【F:f_C_Deploy.cls†L1-L86】
- `DEV_f_pM_EntryLevel.bas` orchestriert den Export sämtlicher versionsrelevanter Daten über `DEV_f_C_VersionControlExport.cls`. Es ruft die einzelnen `bExport…`-Methoden nacheinander auf und stoppt bei Fehlern.【F:DEV_f_pM_EntryLevel.bas†L1-L78】
- `DEV_f_C_VersionControlExport.cls` exportiert Code-Komponenten, Namen, Arbeitsblattmetadaten, Referenzen, Settings-Sheets und definierte Überwachungsbereiche als JSON-Dateien im Projektverzeichnis (z. B. `Names.json`, `WorksheetNames.json`, `VersionControlledRangeContent.json`).【F:DEV_f_C_VersionControlExport.cls†L1-L240】【F:DEV_f_C_VersionControlExport.cls†L240-L437】
- `DEV_f_C_VersionControlRanges.cls` liest Konfigurationen aus `a_wks_VersionControlRanges` (Zeilenstart, Spalten für Name/Defined Name) und baut daraus `DEV_f_C_VersionControlRange`-Objekte. Der Backlog weist auf fehlende Unterstützung für abweichende Arbeitsmappen hin.【F:DEV_f_C_VersionControlRanges.cls†L1-L160】
- Die erzeugten JSON-Dateien geben einen schnellen Überblick über aktuelle Definitionen (z. B. `Names.json`, `WorksheetNames.json`, `VersionControlledRangeContent.json`, `SettingsSheet-*.json`).【F:Names.json†L1-L38】【F:WorksheetNames.json†L1-L40】【F:VersionControlledRangeContent.json†L1-L28】【F:SettingsSheet-a_wks_VersionControlRanges.json†L1-L9】

## 10. Test-Infrastruktur
- `DEV_f_pM_Testing.bas` und `DEV_f_C_UnitTest.cls` legen die Struktur für unit-testbare Prozeduren fest. Tests werden via `DEV_f_p_RegisterUnitTest` gesammelt; eine tatsächliche Ausführungsroutine (`DEV_f_m_RunUnitTests`) ist noch nicht implementiert (TODO-Hinweis).【F:DEV_f_pM_Testing.bas†L1-L46】
- Lower-Level- und Entry-Level-Templates registrieren Tests automatisch, wenn `oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun` gesetzt ist. Für produktiven Einsatz muss die Testlauf-Implementierung ergänzt und ggf. ein UI-Einstiegspunkt geschaffen werden.

## 11. Entwicklungs-Workflows
1. **Neuen Entry-Level-Prozess anlegen:**
   - Prozedur in `a_M_UserInterface` erstellen, die eine Entry-Level-Prozedur in `a_pM_EntryLevel` aufruft.
   - Entry-Level-Prozedur nach Vorlage in `f_pM_TemplatesCore` erstellen, passende Start-/Endmodi wählen und Fehlertexte definieren.【F:f_pM_TemplatesCore.bas†L16-L107】
2. **Lower-Level-Funktionen implementieren:**
   - Basierend auf `b_f_p_TemplateLowerLevel` definieren; Error-Handling-Blöcke beibehalten und TODO-Platzhalter ersetzen.【F:f_pM_TemplatesCore.bas†L108-L188】
3. **Einstellungen pflegen:**
   - Neue Settings in den vorgesehenen Arbeitsblättern eintragen (`a_wks_Settings`, `af_wks_Settings`) und mittels DEV-Helfern (`DEV_f_M_UserInterface`) Namen vergeben.【F:DEV_f_M_UserInterface.bas†L1-L38】
4. **Fehlerhooks erweitern:**
   - App-spezifische Enum-Werte und Logik in `af_pM_ErrorHandling.bas` ergänzen, um detailreichere Meldungen oder automatische Reaktionen zu ermöglichen.【F:af_pM_ErrorHandling.bas†L20-L66】
5. **Versionskontrolle & Deployment:**
   - Bei Bedarf `DEV_f_p_ExportDataForVersionControl` ausführen, um JSON-Dumps und Modul-Exports zu aktualisieren.【F:DEV_f_pM_EntryLevel.bas†L21-L78】
   - Vor Produktivsetzung `f_C_Deploy.bSaveAsProdAndRemoveDEVModules` aufrufen, um DEV-Komponenten zu entfernen.【F:f_C_Deploy.cls†L1-L86】

## 12. Bekannte TODOs und offene Punkte
- **Templates:** Implementierung der Beispielkörper in `f_p_TemplateSubEntryLevel` und `b_f_p_TemplateLowerLevel` steht aus.【F:f_pM_TemplatesCore.bas†L44-L120】
- **RangeArrayProcessor:** Konstruktor, Zeilen-/Spalten-Dictionaries und Namensunterstützung fehlen noch.【F:f_C_RangeArrayProcessor.cls†L23-L46】
- **Worksheet-Klasse:** Ereignisbehandlung für `f_C_Wks` soll laut Backlog erweitert werden.【F:f_C_Wks.cls†L20-L205】
- **Unit-Testing:** `DEV_f_m_RunUnitTests` enthält nur ein TODO und muss implementiert werden.【F:DEV_f_pM_Testing.bas†L34-L46】
- **Version-Control-Ranges:** Unterstützung für Arbeitsmappen außerhalb von `ThisWorkbook` ist angemerkt.【F:DEV_f_C_VersionControlRanges.cls†L19-L30】
- **Fehler-Hooks:** In `af_pM_ErrorHandling.bas` markieren Platzhalter (`>>>>>>> Add your code here`) fehlende Implementierungen für Entry-Level- und Lower-Level-Hooks.【F:af_pM_ErrorHandling.bas†L46-L80】

## 13. Anhänge und Referenzdateien
- `Names.json`, `WorksheetNames.json`, `References.json`, `SettingsSheet-*.json` und `VersionControlledRangeContent.json` dienen als Exportartefakte für Versionskontrolle und Nachvollziehbarkeit aktueller Workbook-Strukturen.【F:Names.json†L1-L38】【F:WorksheetNames.json†L1-L40】【F:SettingsSheet-a_wks_Settings.json†L1-L13】【F:VersionControlledRangeContent.json†L1-L28】
- Weitere DEV-Artefakte (z. B. Test-Canvas-Blätter) befinden sich als `.cls`-Exporte im Projektverzeichnis und werden von `DEV_f_C_VersionControlExport` aktualisiert.

---
Dieses AutoManual fasst alle relevanten Kernmodule, Erweiterungspunkte und offenen Arbeiten im Hauptverzeichnis des Repositories zusammen und dient als schnelle Einstiegshilfe für Entwicklerinnen und Entwickler, die auf Basis von Flow Framework 2 produktiv arbeiten möchten.
