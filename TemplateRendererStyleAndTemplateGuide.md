Note: this is a legacy document referring to the Template Renderer prior to integration into the framework - thus it is not fully correct any longer

# Ziel
Klare Schritt‑für‑Schritt‑Anleitung, um
1) das **Stil‑Sheet `_meta`** (benannter Bereich `Styles`) einzurichten und
2) die **Template‑Sheets** für beide Renderer (einfacher Renderer & Blocks/Lanes‑Renderer) korrekt aufzusetzen.

---

## Voraussetzungen
- Arbeitsmappe enthält (oder bekommt):
  - Blatt **`_meta`** für Styles
  - Blatt **`Template1`** & **`Output1`** (erster Renderer)
  - Blatt **`Template2`** & **`Output2`** (zweiter Renderer)
- VBA‑Module sind eingebunden:
  - `modStyles` (neue Version mit BorderSpec/BorderWeight)
  - Renderer 1 (einfach, zeilenbasiert) – nutzt `ApplyStylesToRange`
  - Renderer 2 (Blocks/Lanes mit `fix_`, `rep_`, `rel_`) – nutzt `ApplyStylesRow2`
- Jeder Renderer ruft **vor** dem Rendern auf: `EnsureStylesFromMeta "_meta"`.

---

# 1) Stil‑Sheet `_meta` einrichten

## 1.1 Benannten Bereich `Styles` anlegen
1) Gehe auf Blatt **`_meta`**.
2) Lege eine Tabelle mit **genau dieser Kopfzeile** in Zeile 1 an (Reihenfolge egal, Schreibweise nicht):
```
Token | NumberFormat | HAlign | VAlign | Wrap | Indent | FontName | FontSize | Bold | Italic | FontColor | FillColor | BorderSpec | BorderWeight
```
3) Markiere die **ganze Tabelle inkl. Header** und vergib den **Namen**: `Styles` (Formel > Namensmanager).

## 1.2 Beispiel‑Zeilen (Start‑Set)
Trage exemplarisch folgende Styles ein (weitere jederzeit ergänzbar):

| Token     | NumberFormat              | HAlign | VAlign | Wrap  | Indent | FontName | FontSize | Bold | Italic | FontColor | FillColor | BorderSpec                    | BorderWeight |
|-----------|---------------------------|--------|--------|-------|--------|----------|----------|------|--------|-----------|-----------|-------------------------------|--------------|
| Body      | General                   | Left   | Center | FALSE | 0      | Calibri  | 11       | FALSE| FALSE  |           |           |                               |              |
| H1        | General                   | Left   | Top    | FALSE | 0      | Calibri  | 16       | TRUE | FALSE  |           |           | BOTTOM                        | THIN         |
| TH        | General                   | Center | Center | TRUE  | 0      | Calibri  | 11       | TRUE | FALSE  |           |           | BOTTOM                        | MEDIUM       |
| Money     | [$€-de-DE] #,##0.00      | Right  | Center | FALSE | 0      | Calibri  | 11       | FALSE| FALSE  |           |           |                               |              |
| Date      | jjjj-mm-tt               | Left   | Center | FALSE | 0      | Calibri  | 11       | FALSE| FALSE  |           |           |                               |              |
| TotalLine | [$€-de-DE] #,##0.00      | Right  | Center | TRUE  | 0      | Calibri  | 11       | TRUE | FALSE  |           |           | TOP                           | MEDIUM       |
| Box       | General                   | Left   | Center | FALSE | 0      | Calibri  | 11       | FALSE| FALSE  |           |           | OUTLINE                       | THIN         |
| Grid      | General                   | Left   | Center | FALSE | 0      | Calibri  | 11       | FALSE| FALSE  |           |           | OUTLINE; INSIDEH; INSIDEV     | THIN         |

**Erläuterungen:**
- **BorderSpec** (Mehrfach möglich): `OUTLINE`, `TOP`, `BOTTOM`, `LEFT`, `RIGHT`, `INSIDEH`, `INSIDEV`.
- **BorderWeight**: `THIN`, `MEDIUM`, `THICK` (oder leer).
- **FontColor/FillColor**: erlaubt `#RRGGBB`, `R,G,B` oder numerischer VBA‑Farbwert.

> Nach Anpassen der Tabelle: Im Renderer vor dem Output **einmal** `EnsureStylesFromMeta "_meta"` aufrufen.

---

# 2) Templates für den **ersten (einfachen) Renderer**
**Charakter:** Zeilenweises Rendern; eine Repeater‑Zeile; Zellen mit Mischtext+Platzhaltern; Styles per Kommentar `style:<Token>`.

## 2.1 Blätter
- **`Template1`**: Vorlage
- **`Output1`**: Ziel

## 2.2 Platzhalter‑Syntax
- Statische Zelle mit Platzhalter: `Rechnung Nr.: {{Invoice.Number}}`
- Repeater‑Platzhalter: `{{Items[i].Name}}`, `{{Items[i].Qty}}`, `{{Items[i].Price}}`, `{{Items[i].Total}}`

## 2.3 Repeater‑Zeile anlegen
1) Lege **eine** Musterzeile an (z. B. Zeile 6):
   - `A6: {{Items[i].Name}}`
   - `B6: {{Items[i].Qty}}`
   - `C6: {{Items[i].Price}}`
   - `D6: {{Items[i].Total}}`
2) **Benannter Bereich** über die ganze Zeile: `rep_Items` (genau so).

> Im Standard erwartet der einfache Renderer die Repeater‑Daten unter **Key `"Items"`**.

## 2.4 Styles setzen (Kommentare)
- Pro Zelle, die formatiert werden soll, Kommentar: `style:<Token>`
  - Bsp.: Überschrift `H1`, Kopf `TH`, Geld `Money`, Datum `Date`, Summe `TotalLine`.
- Die Repeater‑Zeile bekommt die Style‑Kommentare in allen Spalten, die wiederholt werden.

## 2.5 Typischer Aufbau (Minimal)
- `A1 (H1)`: `Rechnung Nr.: {{Invoice.Number}}`
- `C1 (H1)`: `Datum: {{Invoice.Date}}`
- `A3 (Body)`: `Kunde: {{Customer.Name}}`
- `A5:D5 (TH)`: Spaltenköpfe
- `rep_Items` in Zeile 6 (siehe oben)
- `D8 (TotalLine)`: `{{Totals.Sum}}`

## 2.6 Render‑Ablauf (Kurz)
- Vorher: `EnsureStylesFromMeta "_meta"`.
- Renderer schreibt statische Zeilen + expandiert `rep_Items` (n Einträge).
- Pro Zielzelle: `.Style = Token` und **`ApplyBordersForToken`** aus `modStyles`.

---

# 3) Templates für den **zweiten Renderer (Blocks/Lanes)**
**Charakter:** Mehrere **Blöcke**; je Block **Lanes**: `fix_`, `rep_`, **`rel_`**; dynamische Höhen; `padAfter` (Leerzeilen) je Lane.

## 3.1 Blätter
- **`Template2`**: Vorlage
- **`Output2`**: Ziel

## 3.2 Named‑Range‑Konventionen
- **Block:** `blk_<BlockKey>` → zusammenhängender Bereich für den Block.
- **Statische Lane:** `fix_<BlockKey>_<LaneKey>` – liegt vollständig **in** `blk_…`.
- **Repeater‑Lane:** `rep_<BlockKey>_<LaneKey>` – Musterzeile(n), vollständig **in** `blk_…`.
- **Relative Lane (verschiebbar):** `rel_<BlockKey>_<LaneKey>` – steht an Template‑Position, wird aber **nach unten geschoben**, wenn darüber etwas expandiert.

**Regeln:**
- Jede `fix_`/`rep_`/`rel_`‑Range muss **vollständig innerhalb** ihres `blk_…` liegen.
- **LaneKey** muss den Daten‑Key im Kontext treffen. Beispiel: `rep_Panel_Items` ⇒ LaneKey = `Items` ⇒ Daten unter `ctx("repeaters")("Items")`.

## 3.3 Styles & Metadaten (Kommentare)
- Kommentare in Zellen: `style:<Token>`
- Optional pro `rep_`/`rel_` **padAfter** (Leerzeilen nach der Lane):
  - In **erster Zelle der Lane**: `style:<Token>; padAfter:1`

## 3.4 Typisches Beispiel
**Panel‑Block (fix links, rep rechts, rel Footer):**
- `blk_Panel = A3:F5`
- `fix_Panel_Left = A3:C5`
  - `A3 (Body)`: `Kunde: {{Customer.Name}}`
  - `A4 (Body)`: `Ort: {{Customer.City}}`
  - `A5 (Body)`: `Land: {{Customer.Country}}`
- `rep_Panel_Items = D3:F3` (Musterzeile)
  - `D3 (Body)`: `{{Items[i].Date}}`
  - `E3 (Body)`: `{{Items[i].Ref}}`
  - `F3 (Money)`: `{{Items[i].Amount}}`
  - Kommentar in D3: `style:Body; padAfter:1` (→ **eine Leerzeile** unter den Items)
- `rel_Panel_Footer = A6:C6` (z. B. `Hinweis …`), Kommentar `style:Body`

**Invoice‑Block (klassischer Repeater):**
- `blk_Invoice2 = A8:D13`
- `fix_Invoice2_Header = A8:D9` (H1/TH)
- `rep_Invoice2_Items = A10:D10`
  - `A10 (Body)`: `{{Items[i].Name}}`
  - `B10 (Body)`: `{{Items[i].Qty}}`
  - `C10 (Money)`: `{{Items[i].Price}}`
  - `D10 (Money)`: `{{Items[i].Total}}`
- `fix_Invoice2_Footer = A12:D13` (z. B. `D12 (TotalLine)`: `{{Totals.Sum}}`)

**Wirkprinzip:**
- Reihenfolge im Output bestimmt sich durch **tatsächliche Blockhöhe**.
- `rep_` expandiert nach unten; `rel_` wird **unter** bereits Geschriebenes geschoben (`max(TemplateTopRel, currentMaxBottom+1)`).
- `padAfter` fügt **Leerzeilen** hinter der Lane ein.

---

# 4) Platzhalter‑Regeln (beide Renderer)
- Syntax: `{{Key}}` – nur der Platzhalter wird ersetzt, **statische Textteile** bleiben erhalten.
- In Repeatern: `{{Items[i].Feld}}` (der Index `i` wird durch den Renderer je Zeile ersetzt).
- Totals/Header: frei benennbar (z. B. `Totals.Sum`, `Invoice.Number`, `Customer.City`).

---

# 5) Start & Smoke‑Tests
1) `_meta` füllen → **Makro** `EnsureStylesFromMeta "_meta"` wird beim Rendererstart aufgerufen.
2) `Template1`/`Template2` nach obigen Regeln befüllen & benennen.
3) **Render starten**:
   - Einfacher Renderer → nach `Output1`
   - Blocks/Lanes‑Renderer → nach `Output2`

**Prüfen:**
- Werden Styles (Ausrichtung, Font, NumberFormat) angewandt?
- Stimmen Rahmen mit `BorderSpec`/`BorderWeight` (z. B. `BOTTOM` dünn vs. mittel)?
- Werden Repeaterzeilen korrekt vervielfacht; erscheinen `rel_`‑Lanes **unter** Expandiertem?

---

# 6) Troubleshooting (kurz)
- **Nur linker Rahmen sichtbar:** Stelle sicher, dass `ApplyBordersForToken` im Renderer pro **Zielzelle** aufgerufen wird und `ClearManagedBorders` vorher Kanten löscht. `BorderSpec` muss exakt `TOP/BOTTOM/...` enthalten.
- **Keine Items im Repeater:** `LaneKey` prüfen (z. B. `rep_Panel_Items` ⇒ Key = `Items`) und Daten unter `ctx("repeaters")("Items")` bereitstellen.
- **`IIf`‑Falle in VBA:** `IIf` evaluiert beide Zweige → nie `.Count` o.Ä. im „sonst“-Zweig; lieber vorher per `If coll Is Nothing Then n=0 Else n=coll.Count`.
- **Objektzuweisungen:** Für Collections/Dictionarys immer **`Set`** verwenden (z. B. `Set ctx("repeaters")("Items") = col`).
- **Named Ranges:** `fix_`/`rep_`/`rel_` müssen vollständig **innerhalb** ihres `blk_` liegen.

---

# 7) Erweiterungen (optional)
- Weitere Border‑Profile: z. B. `HeaderBox = OUTLINE; BOTTOM; INSIDEV` mit `MEDIUM`.
- Zusätzliche Template‑Metadaten: eigene Kommentar‑Keys (z. B. `visibleIf:…`).
- Post‑Processing: Gesamtranges als Block umranden: `ApplyBordersForToken Range("A5:D20"), "Grid"`.

---

## Kurzfazit
- `_meta!Styles` steuert **alle** Formatdetails + Rahmen.
- Renderer 1: schnell & simpel (eine Repeater‑Zeile).
- Renderer 2: flexibel (mehrere Blöcke, `fix`/`rep`/`rel`, `padAfter`).
- Kommentare `style:<Token>` + optionale `padAfter:<n>` geben Dir volle Layout‑Kontrolle direkt im Template.

