# Roadmap & Sprint-Ready User Stories (Flow Framework 2)

## Zielbild
Diese Roadmap beschreibt die Einführung von sieben priorisierten Funktionalitäten zur Reifung von **Flow Framework 2** in Richtung: höhere Qualität, bessere Wartbarkeit, zuverlässigere Releases und schnellere Implementierung in Folgeprojekten.

---

## 1) Priorisierte Roadmap

## Planungsannahmen
- Team: 1–2 VBA-Entwickler + 1 Product Owner (Teilzeit).
- Sprintlänge: 2 Wochen.
- Kapazität: ca. 18–24 Story Points pro Sprint.
- Reihenfolge ist auf schnelle Risikoreduktion optimiert (erst Qualitätssicherung, dann Erweiterbarkeit und Tooling).

## Release-Backlog (high level)

| Reihenfolge | Feature | Ziel | Aufwand (grob) | Abhängigkeiten |
|---|---|---|---:|---|
| 1 | DEV Test Framework vervollständigen | Verlässliche Regressionstests | M (13 SP) | keine |
| 2 | Strikter Config/JSON-Validierungsmodus | Parsebare Artefakte & Automatisierung | M (8 SP) | #1 (Tests als Guardrail) |
| 3 | Erweiterte Error-Observability | Bessere Analyse von Laufzeitfehlern | M (8 SP) | #1 |
| 4 | Konfigurierbare Deployment-Pipeline | Reproduzierbare, sichere Releases | L (13 SP) | #1, #2, #3 |
| 5 | Hook-/Plugin-System (Lifecycle) | Saubere Erweiterbarkeit | L (13 SP) | #3 |
| 6 | Renderer-Tooling (Lint/Preview/Diagnostics) | Schnellere Template-Entwicklung | L (13 SP) | #2, #5 (optional) |
| 7 | Migration-/Upgrade-Assistent | Stabiler Versionswechsel | L (13 SP) | #2, #4 |

## Sprint-Vorschlag (4 Sprints)

### Sprint 1 (Foundation Quality)
- Feature #1 DEV Test Framework (vollständig)
- Feature #2 JSON/Config Strict Mode (MVP)

### Sprint 2 (Runtime Reliability)
- Feature #3 Error Observability
- Feature #4 Deployment-Pipeline (MVP mit Build-Profilen + Preflight)

### Sprint 3 (Extensibility & UX für Entwickler)
- Feature #5 Hook-/Plugin-System
- Feature #6 Renderer-Linter + Diagnostics (MVP)

### Sprint 4 (Governance & Long-Term Maintainability)
- Feature #4 Deployment-Pipeline (Härtung/Erweiterung)
- Feature #6 Renderer Preview/Dry-Run
- Feature #7 Migration-/Upgrade-Assistent

---

## 2) Sprint-Ready User Stories je Funktionalität

> Format: Story, Business Value, Scope, Nicht-Ziele, Akzeptanzkriterien (Given/When/Then), Tasks, Abhängigkeiten, Risiken, Testfälle, Definition of Ready, Definition of Done, Schätzung.

## Feature 1: DEV Test Framework vervollständigen

### Story 1.1 – Test Runner ausführbar machen
**Als** Framework-Entwickler  
**möchte ich** alle registrierten Unit Tests per Entry-Level-Prozedur starten können  
**damit** ich Änderungen schnell gegen Regressionen prüfen kann.

**Business Value:** Hoch (Basis für alle Folgefeatures)

**In Scope**
- Implementierung `DEV_f_m_RunUnitTests` inkl. Discovery/Registrierung von Testsuiten.
- Konsolen-/Immediate-Output plus kompaktes Ergebnisobjekt (Pass/Fail/Skipped).

**Out of Scope**
- CI-Server-Integration.

**Akzeptanzkriterien**
1. **Given** mindestens zwei Testsuiten mit je >1 Test, **when** Runner startet, **then** werden alle Tests deterministisch ausgeführt.
2. **Given** ein fehlschlagender Test, **when** Run endet, **then** enthält Summary Anzahl Passed/Failed und Fehlertexte.
3. **Given** keine Tests registriert, **when** Run startet, **then** beendet der Runner kontrolliert mit „0 tests executed“.

**Tasks (technisch)**
- Testregistrierung (Collection/Dictionary) festlegen.
- Runner-Schleife + Ergebnisaggregation.
- Standardisierte Reporter-Ausgabe.

**Abhängigkeiten:** keine  
**Risiken:** Unterschiedliches Verhalten bei nicht initialisierten Objekten in VBA.  
**Testfälle:** 1x alle grün, 1x absichtlicher Fail, 1x leere Registry.

**Definition of Ready (DoR)**
- Namenskonvention für Tests dokumentiert.
- Mindestens 2 Pilot-Testmodule vorhanden.

**Definition of Done (DoD)**
- Runner liefert reproduzierbare Resultate.
- Ergebniszusammenfassung maschinenlesbar (mind. strukturierte Textform).
- Dokumentation „How to run tests“ ergänzt.

**Schätzung:** 5 SP

---

### Story 1.2 – Assertions & Test-Hilfsfunktionen bereitstellen
**Als** Entwickler  
**möchte ich** robuste Assert-Funktionen nutzen  
**damit** Tests klar, kurz und wartbar bleiben.

**Business Value:** Hoch

**In Scope**
- `AssertEqual`, `AssertTrue`, `AssertFalse`, `AssertThrows`, `AssertNotNothing`.
- Einheitliches Fehlformat (Expected/Actual/Context).

**Out of Scope**
- Snapshot-Testing.

**Akzeptanzkriterien**
1. Assertion-Fehler zeigen Expected/Actual und Testnamen.
2. `AssertThrows` besteht nur, wenn der erwartete Fehler auftritt.
3. Assertions sind in mindestens 10 Referenztests verwendet.

**Tasks**
- Assertion-Modul erstellen.
- Beispieltests migrieren.
- Failure-Message-Format standardisieren.

**Abhängigkeiten:** Story 1.1  
**Risiken:** Fehlertypen in VBA teils uneinheitlich.  
**Testfälle:** Pro Assertion mind. 1 Pass + 1 Fail.

**DoR:** Fehlermeldungsformat abgestimmt.  
**DoD:** Assertions dokumentiert + in Pilot-Suite produktiv genutzt.

**Schätzung:** 3 SP

---

### Story 1.3 – Testergebnisse exportieren
**Als** Teammitglied  
**möchte ich** Testresultate als JSON/Sheet exportieren  
**damit** ich Läufe historisieren und vergleichen kann.

**Business Value:** Mittel

**In Scope**
- JSON-Export (`DEV-TestReport.json`) und optionales Ergebnis-Sheet.
- Zeitstempel, Dauer, Counters, Fehlerdetails.

**Akzeptanzkriterien**
1. Exportdatei wird pro Lauf überschrieben oder versioniert (konfigurierbar).
2. JSON ist strikt parsebar.
3. Fehlgeschlagene Tests enthalten Stack/Call-Kontext soweit verfügbar.

**Schätzung:** 5 SP

---

## Feature 2: Strikter Config/JSON-Validierungsmodus

### Story 2.1 – Strict JSON Exportmodus
**Als** Entwickler  
**möchte ich** einen strikt validen JSON-Modus für Exportartefakte  
**damit** externe Tools die Dateien zuverlässig parsen.

**Business Value:** Hoch

**In Scope**
- Schalter `ExportMode = "lenient" | "strict"`.
- Strict-Ausgabe für `SettingsSheet-*.json`, `Names.json`, `WorksheetNames.json`, `References.json`.

**Akzeptanzkriterien**
1. Im Strict-Modus enthalten Dateien keine trailing commas.
2. Alle Strict-Dateien bestehen einen JSON-Parser-Check.
3. Lenient-Modus bleibt rückwärtskompatibel.

**Tasks**
- Serializer-Pfad aufteilen (lenient/strict).
- Validierungshook nach Export.
- Rückwärtskompatibilitäts-Testfälle.

**Abhängigkeiten:** Feature 1 hilfreich

**Schätzung:** 5 SP

---

### Story 2.2 – Config-Validierung mit Fehlerbericht
**Als** Product Owner  
**möchte ich** beim Export einen strukturierten Validierungsbericht  
**damit** Konfigurationsprobleme früh sichtbar sind.

**Business Value:** Mittel/Hoch

**In Scope**
- Prüfungen: Pflichtfelder, Datentypen, Duplikate, leere Keys.
- Report als JSON + kurze UI-Zusammenfassung.

**Akzeptanzkriterien**
1. Validierungsfehler blockieren Strict-Export (konfigurierbar).
2. Report enthält Fundstelle (Datei/Sheet/Name/Feld).
3. Warnungen und Fehler sind getrennt klassifiziert.

**Schätzung:** 3 SP

---

## Feature 3: Erweiterte Error-Observability

### Story 3.1 – Korrelation & Severity im Error-Objekt
**Als** Support-Entwickler  
**möchte ich** jede Fehlerkette über Korrelations-ID und Severity nachverfolgen  
**damit** ich Ursachen schneller identifizieren kann.

**Business Value:** Hoch

**In Scope**
- Ergänzung Fehlerdatenmodell: `CorrelationId`, `Severity`, `TimestampUtc`, `ProcedureName`.
- Automatische Vergabe pro Entry-Level-Lauf.

**Akzeptanzkriterien**
1. Jeder registrierte Fehler enthält CorrelationId.
2. Severity-Werte sind konsistent (`INFO|WARN|ERROR|FATAL`).
3. Fehlerkette bleibt über Aufruftiefe hinweg korreliert.

**Schätzung:** 5 SP

---

### Story 3.2 – Fehlerjournal & Dashboard
**Als** Fachanwender mit Maintenance-Rechten  
**möchte ich** ein aktuelles Fehlerjournal sehen  
**damit** ich wiederkehrende Probleme erkenne.

**Business Value:** Mittel

**In Scope**
- Persistenz in Log-Sheet oder Datei.
- Einfaches Dashboard (Top Fehler, letzter Lauf, Häufigkeit).

**Akzeptanzkriterien**
1. Neue Fehler erscheinen mit Zeit, CorrelationId, Kurztext.
2. Dashboard aktualisiert sich nach jedem Lauf.
3. Maintenance-Only Sichtbarkeit ist konfigurierbar.

**Schätzung:** 3 SP

---

## Feature 4: Konfigurierbare Deployment-Pipeline

### Story 4.1 – Build-Profile (Debug/QA/Prod)
**Als** Release-Verantwortlicher  
**möchte ich** vordefinierte Build-Profile nutzen  
**damit** Deployments reproduzierbar und nachvollziehbar sind.

**Business Value:** Hoch

**In Scope**
- Profilkonfiguration (Flags, DEV-Entfernung, Logging-Level, Testpflicht).
- Entry-Level: `DeployWithProfile(profileName)`.

**Akzeptanzkriterien**
1. Profile sind zentral konfigurierbar.
2. Deploy bricht ab, wenn Pflichtbedingungen des Profils nicht erfüllt sind.
3. Build-Manifest (Version, Profil, Zeit, Ergebnis) wird erzeugt.

**Schätzung:** 5 SP

---

### Story 4.2 – Preflight- und Post-Deploy-Checks
**Als** Release-Verantwortlicher  
**möchte ich** automatisierte Vor- und Nachprüfungen  
**damit** fehlerhafte Builds nicht ausgeliefert werden.

**In Scope**
- Preflight: Tests grün, Strict-Export valide, erforderliche Settings vorhanden.
- Postflight: Workbook öffnet, Kern-Entry-Level aufrufbar, DEV-Module entfernt (Prod).

**Akzeptanzkriterien**
1. Ein fehlschlagender Preflight stoppt den Deploy.
2. Postflight erzeugt Checkprotokoll.
3. Ergebnisse werden im Build-Manifest referenziert.

**Schätzung:** 8 SP

---

## Feature 5: Hook-/Plugin-System (Lifecycle)

### Story 5.1 – Lifecycle-Events definieren & registrieren
**Als** App-Entwickler  
**möchte ich** standardisierte Lifecycle-Hooks registrieren  
**damit** ich Erweiterungen ohne Core-Fork einhängen kann.

**Business Value:** Hoch

**In Scope**
- Hook-Punkte: `BeforeStart`, `AfterInit`, `BeforeRender`, `OnErrorHandled`, `BeforeDeploy`, `AfterDeploy`.
- Registry + Ausführungsreihenfolge (Priority).

**Akzeptanzkriterien**
1. Hooks können pro Event mehrfach registriert werden.
2. Reihenfolge ist deterministisch (Priority, dann Name).
3. Fehler in Hook A verhindern nicht zwingend Hook B (Policy-konfigurierbar).

**Schätzung:** 8 SP

---

### Story 5.2 – Plugin-Konfiguration per Settings
**Als** Product Owner  
**möchte ich** Plugins per Settings ein-/ausschalten  
**damit** ich Verhalten ohne Codeänderung steuern kann.

**Akzeptanzkriterien**
1. Deaktivierte Plugins werden nicht geladen.
2. Fehlkonfigurationen erzeugen klare Warnungen.
3. Aktive Plugin-Liste ist zur Laufzeit abrufbar.

**Schätzung:** 5 SP

---

## Feature 6: Renderer-Tooling (Lint/Preview/Diagnostics)

### Story 6.1 – Template-Linter
**Als** Template-Entwickler  
**möchte ich** Vorabprüfungen für Templates  
**damit** ich Struktur- und Tokenfehler vor dem Rendern finde.

**Business Value:** Hoch

**In Scope**
- Lint-Regeln für `blk_`, `fix_`, `rep_`, `rel_`, Platzhalter, Style-Tokens.
- Ergebnisliste mit Severity + Zellreferenz.

**Akzeptanzkriterien**
1. Fehlende/inkonsistente Blöcke werden erkannt.
2. Unaufgelöste Platzhalter werden als Fehler/Warnung markiert.
3. Lint-Report ist exportierbar.

**Schätzung:** 8 SP

---

### Story 6.2 – Dry-Run Preview mit Diagnostics
**Als** Template-Entwickler  
**möchte ich** eine Render-Vorschau ohne persistente Änderungen  
**damit** ich Iterationen beschleunige.

**Akzeptanzkriterien**
1. Preview kann auf Kopie/temporärem Sheet laufen.
2. Diagnostics zeigen expandierte Blöcke, Laufzeiten, angewandte Styles.
3. Abbruch bei kritischen Fehlern ist möglich.

**Schätzung:** 5 SP

---

## Feature 7: Migration-/Upgrade-Assistent

### Story 7.1 – Strukturelle Bestandsaufnahme (Audit)
**Als** Entwickler  
**möchte ich** vor einem Upgrade einen automatischen Struktur-Audit  
**damit** ich Inkompatibilitäten früh sehe.

**Business Value:** Hoch

**In Scope**
- Prüfung erwarteter Names, Sheets, Settings, Versionsmarker.
- Audit-Report mit Must-Fix/Should-Fix.

**Akzeptanzkriterien**
1. Audit listet fehlende Elemente inkl. Priorität.
2. Report enthält empfohlene Migrationsschritte.
3. Audit kann standalone ohne Deploy laufen.

**Schätzung:** 5 SP

---

### Story 7.2 – Geführte Migration mit Rollback-Punkt
**Als** Release-Verantwortlicher  
**möchte ich** ein geführtes Upgrade inkl. Backup/Rollback  
**damit** ich Risiko bei Versionswechsel minimiere.

**In Scope**
- Schrittweise Migration mit Checkpoints.
- Backup vor Start, Rollback bei Fehler.

**Akzeptanzkriterien**
1. Vor Migration wird ein Backup-Artefakt erstellt.
2. Bei Fehler in Schritt N erfolgt Rollback auf letzten stabilen Stand.
3. Abschlussbericht dokumentiert alle Schritte und Ergebnisse.

**Schätzung:** 8 SP

---

## 3) Übergreifende Sprint-Readiness-Checkliste

Eine Story gilt erst dann als sprint ready, wenn alle Punkte erfüllt sind:

- [ ] Fachlicher Nutzen in 1–2 Sätzen beschrieben.
- [ ] Klare In-/Out-of-Scope Abgrenzung.
- [ ] Akzeptanzkriterien testbar (Given/When/Then oder äquivalent).
- [ ] Technische Tasks in umsetzbare Schritte geschnitten (max. 1–2 Tage pro Task).
- [ ] Abhängigkeiten und Risiken dokumentiert.
- [ ] Testfälle (mind. Happy Path + 1 Edge Case + 1 Failure Case) definiert.
- [ ] Messkriterium für Erfolg vorhanden (z. B. Parsing-Rate 100%, Deploy-Fehlerrate sinkt).
- [ ] Schätzung in Story Points vorhanden und teamseitig bestätigt.

---

## 4) Empfohlene Metriken zur Steuerung

- Testabdeckung kritischer Entry-Level-Prozeduren (%).
- Durchschnittliche Zeit bis Fehlerursache identifiziert ist (MTTI).
- Anteil Strict-JSON-konformer Exporte (%).
- Deploy-Erfolgsquote pro Profil (%).
- Anzahl Renderer-Fehler vor/nach Linter-Einführung.
- Upgrade-Erfolgsquote ohne manuellen Eingriff (%).

---

## 5) Vorschlag für initiale Epic-Struktur im Backlog-Tool

- **EPIC A – Quality Foundation**: Feature 1 + 2
- **EPIC B – Runtime & Release Reliability**: Feature 3 + 4
- **EPIC C – Extensibility & Template DX**: Feature 5 + 6
- **EPIC D – Lifecycle Governance**: Feature 7

Damit ist ein umsetzbarer Pfad vorhanden, der erst die technische Basis stabilisiert und dann Erweiterbarkeit + langfristige Wartbarkeit skaliert.
