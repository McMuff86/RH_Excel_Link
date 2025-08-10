### ExcelLink – Korrex-README (Stand: aktuelle Session)

Dieses Dokument beschreibt präzise, was wir implementiert haben, wie das Plug-in funktioniert, welche Entscheidungen/Workarounds nötig waren, und worauf bei einer Neuimplementierung zu achten ist, damit der Code beim ersten Mal funktioniert.

---

#### Ziel
- Excel-Tabellen in Rhino 8 (Windows) einfügen, aktualisieren und zurückschreiben (ClosedXML; keine OLE/Interop).
- Tabelle wird als Block (Linien + Text) gerendert und mit Excel (Datei/Sheet/Range) verknüpft.
- Optionen: Grid, Alignment, Wrap, Textgröße/Font bzw. DimStyle, Scale, Insert-Punkt.
- Update via Command und FileWatcher (mit Prompt), Relink, sowie ein Command zum Anwenden von Attributen aus UserText.

---

#### Projektaufteilung (wichtig zum Verständnis)
- `ExcelLink.Core`
  - Excel I/O (ClosedXML) und Datenmodell (`TableModel` → `Row` → `Cell`).
  - `ExcelInspector`: Hilfsfunktionen zum Ermitteln von Sheets und UsedRange in A1-Notation.
- `ExcelLink.Plugin`
  - Rhino-spezifisches Rendering, Commands, UI (Eto.Forms).
  - Commands:
    - `ExcelLinkInsert`
    - `ExcelLinkUpdate`
    - `ExcelLinkRelink`
    - `ExcelLinkApplyAttributes`
  - Services:
    - `ExcelLinkWatcherService`: FileSystemWatcher mit Update-Prompt und optionalem Range-Extend.

---

#### UserText-Metadaten (am `InstanceDefinition` + Instanzen)
Die folgenden Schlüssel werden geschrieben und bei jedem Update/Relink/ApplyAttributes beibehalten:
- `ExcelLink_File` – Pfad zur .xlsx
- `ExcelLink_Sheet` – Blattname
- `ExcelLink_Range` – Bereich oder Named Range (A1:D20, BOM, …)
- `ExcelLink_Scale` – Skalierungsfaktor (mm → Model Units)
- `ExcelLink_Align` – 0/1/2 (Left/Center/Right) bzw. „Excel“ aus dem Dialog
- `ExcelLink_ShowGrid` – 1/0
- `ExcelLink_TextMm` – Zahl (mm) oder leer (leer = Höhe aus DimStyle)
- `ExcelLink_Font` – Font-Family oder leer
- `ExcelLink_Wrap` – 1/0
- `ExcelLink_DimStyle` – Name des DimStyle (z. B. „Standard 1:1“)

Wichtig: Keys werden bei Redefinition der Block-Definition wieder auf die neue Definition und neu eingesetzte Instanzen zurückgeschrieben, sodass Update-Ketten stabil bleiben.

---

#### Commands & Verhalten
- `ExcelLinkInsert`
  - Eto-Dialog: Datei wählen, Sheet/Range, Scale, Alignment, Grid, Wrap, Font, Text (mm), DimStyle (Dropdown), Mini-Preview.
  - Danach Insert-Punkt in Rhino auswählen.
  - Diagramm/Optionen werden als UserText geschrieben.

- `ExcelLinkUpdate`
  - Instanz(en) wählen.
  - Liest Excel neu ein, rendert die Definition neu und setzt alle Instanzen mit den ursprünglichen `InstanceXform`s wieder ein.
  - Re-schreibt UserText (Definition und Instanzen).

- `ExcelLinkRelink`
  - Instanz wählen → Dialog → neue Datei/Sheet/Range setzen → zur Aktualisierung `ExcelLinkUpdate` ausführen.

- `ExcelLinkApplyAttributes`
  - Instanz(en) wählen.
  - Liest die `ExcelLink_*` UserText-Keys von der Instanz und schreibt sie auf die zugehörige Definition, anschließend Redefinition wie bei Update.
  - Effekt: Änderungen im Properties-Panel (UserText) lassen sich ohne Dateiwächter anwenden.

---

#### Rendering-Pipeline (kritische Details)
1) Excel → `TableModel` (ClosedXML): Werte, Alignment, Merges (Platzhalter), Größen (Spaltenbreite/Zeilenhöhe in mm), Wrap-Hinweise.
2) Optionen (UserText/Dialog) werden angewandt.
3) Rendering:
   - Grid-Linien als `LineCurve` (optional).
   - Text je Zelle als `TextEntity`.
   - Alignment horizontal: Left/Center/Right; vertikal center (MVP).
   - Wrap (heuristisch) abhängig von Spaltenbreite.
   - DimStyle:
     - Texte werden – wenn DimStyle gesetzt – mit `TextEntity.Create(text, plane, DimensionStyle, wrapped, rectWidth, rotation)` erzeugt.
     - „Height from DimStyle“ aktiv: keine explizite `TextHeight` setzen.
     - Sonst `TextHeight` setzen (DimStyle bleibt zugewiesen, Höhe wird überschrieben).

4) Block-Definition (wichtigster Workaround):
   - Alle Geometrien (Linien + Texte) werden zunächst temporär als Doc-Objekte hinzugefügt (so übernimmt Rhino Annotation/Styles akkurat).
   - Aus den Doc-Objekten werden Geometrie + Attributes dupliziert und als Block-Definition erstellt.
   - Temporäre Objekte werden wieder gelöscht.
   - Instanz wird mit Insert-Transform eingesetzt.

Warum dieser Weg? Ein direktes `Add(definition, geoms, atts)` verliert bei Rhino gelegentlich den DimStyle (fällt auf „Default“). Über Doc-Objekte bleiben Styles zuverlässig erhalten.

---

#### FileWatcher & Range-Extend
- Beobachtet jedes verknüpfte `.xlsx` (pro Definition) und fragt bei Änderungen:
  - „Update now?“ – führt Update aus.
  - Optional: „Extend ranges to used range?“ – wenn größer geworden: Range aktualisieren und neu laden.

---

#### Mini-Preview
- Einfache Bitmap-Vorschau (Eto) im Dialog basierend auf aktueller Datei/Sheet/Range + Grid/Font/Wrap.
- Dient als Orientierung (kein 1:1 Rendering wie Rhino, aber nah genug für Auswahl).

---

#### Wichtige Learnings (für eine Neuimplementierung ohne Iterationen)
1) DimStyle korrekt binden:
   - TextEntity unbedingt per `TextEntity.Create(text, plane, DimensionStyle, wrapped, rectWidth, rotation)` erzeugen, nicht erst Style später zuweisen.
   - „Height from DimStyle“ bedeutet: keine `TextHeight` setzen. Nur bei explizitem mm-Wert setzen.

2) Block-Definition so erstellen, dass Annotation-Eigenschaften sicher ankommen:
   - Temporär als Doc-Objekte einfügen → daraus Geometrie + Attribute duplizieren → Block-Definition bauen → Temp löschen.

3) UserText-Keys konsequent pflegen:
   - Beim Insert, Update, ApplyAttributes und Relink ALLE Keys (inkl. `ExcelLink_DimStyle`) auf Definition und Instanzen schreiben.

4) Update darf Instanzen nicht „verlieren“:
   - Vor Redefinition alle Instanz-Transforms sammeln, Instanzen löschen, Definition neu aufbauen, Instanzen mit selben XForms neu einfügen.

5) Keine OLE/Interop:
   - Ausschließlich ClosedXML für Lesen/Schreiben (stabil, keine Busy-Prompts).

6) UI-Threads beachten:
   - Schweres außerhalb des UI-Threads. Nur UI via `RhinoApp.InvokeOnUiThread`.

---

#### Bekannte Grenzen / ToDos (MVP)
- Merged Cells: Darstellung als einzelne TextEntity (vereinfachter Umgang) – weiter ausbauen.
- Wrap: heuristisch; später besserer Textumbruch.
- Vertikale Ausrichtung: aktuell Middle; später Top/Bottom.
- Panel: DimStyle- und Optionssteuerung direkt im Panel nachrüsten (Apply-Button vorhanden als Command).

---

#### Build & Installation
- .NET 7, C# 11, Rhino 8, Referenzen: `RhinoCommon.dll`, `Rhino.UI.dll`, `Eto.dll` aus `C:\Program Files\Rhino 8\System`.
- `ExcelLink.Plugin` erzeugt nach Build `ExcelLink.rhp` (Post-Build Copy). Rhino muss beim Build geschlossen sein (Dateisperre!).
- Laden in Rhino: Tools → Options → Plug-ins → Install… → `ExcelLink.rhp` wählen.

---

#### Nutzung (Kurzablauf)
1) `ExcelLinkInsert` → Datei/Sheet/Range/Optionen → Einfügepunkt.
2) Datei ändern → FileWatcher fragt → Yes / Extend → Tabelle aktualisiert sich.
3) `ExcelLinkUpdate` (manuell) → Instanzen wählen → aktualisiert Definition.
4) `ExcelLinkRelink` → Instanz wählen → neue Quelle setzen → `ExcelLinkUpdate`.
5) `ExcelLinkApplyAttributes` → Instanz wählen → UserText-Optionen auf Definition anwenden und rendern.

---

#### Troubleshooting
- Texte zeigen „Default“ statt gewähltem DimStyle:
  - Sicherstellen, dass `TextEntity.Create(..., DimensionStyle, ...)` verwendet wird und beim Blockbau der Doc-Objekt-Workaround genutzt wird.
- `.rhp` lässt sich nicht kopieren: Rhino läuft noch. Plug-in entladen oder Rhino schließen.
- UserText fehlt nach Update: Prüfe, ob Keys in `UpdateAndRelink.cs` konsequent auf Definition + Instanzen gesetzt werden.

---

#### Kurzes Changelog (Session)
- Insert-Dialog (Grid/Alignment/Scale/Font/Text mm/Wrap/DimStyle, Mini-Preview).
- Rendering: Alignment, Grid, heuristisches Wrap.
- DimStyle-Bindung via `TextEntity.Create` + Blockdefinition über temporäre Doc-Objekte.
- Commands: Update, Relink, ApplyAttributes.
- FileWatcher mit Prompt und Range-Extend.
- Persistente UserText-Keys; ApplyAttributes/Update/Watcher schreiben Werte zurück.


