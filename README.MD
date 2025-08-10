### ExcelLink — Rhino 8 C# Plugin (RhinoCommon, .NET 7)

ExcelLink lets you insert, update, and roundtrip Excel tables inside Rhino. Tables are rendered as blocks (lines + text) and linked to an Excel range (file path + sheet + range/named range). No COM/OLE; Excel I/O is handled via ClosedXML.

---

#### Key Principles
- **No COM/OLE-Interop**: Use ClosedXML (NuGet) to read/write `.xlsx` for stability and to avoid OLE busy prompts.
- **UI with Eto.Forms**: Rhino-friendly UI; avoid WPF/WinForms specifics.
- **Responsive UI**: Long operations off the UI thread; use `RhinoApp.InvokeOnUiThread` only for UI work.
- **Undo safety**: Batch geometry changes within a single Undo record.
- **Blocks for sync**: Use block definitions so model/layout instances stay in sync.
- **Metadata on blocks**: Persist path, sheet, range, and options in `UserDictionary`.
- **File watcher**: Monitor source files and prompt to update.
- **Cross-platform awareness**: Excel I/O is Windows-agnostic; Rhino-specific parts follow Rhino 8 conventions.
- **Robust errors**: Clear error handling with messages logged to the Rhino command line.

---

#### Commands
- **ExcelLinkInsert**
  - Dialog to choose `.xlsx`, sheet, range or named range; target (model or existing layout page), scale (mm), text style/height, and options (grid lines on/off, text alignment, borders).
  - Read cells via ClosedXML and materialize a rendered table: rectangles/lines for grid, `TextEntity` per cell.
  - Create block definition `ExcelLink_<hash>` and insert a block instance.
  - Store link metadata on the block (path, sheet, range, formatting options).

- **ExcelLinkUpdate**
  - For selected ExcelLink blocks: reload Excel, rebuild the block definition; instances keep their placement.
  - Includes option to “Update All”.

- **ExcelLinkPanel** (modeless)
  - Eto grid that displays the linked table and allows basic editing (text/number; no formula support).
  - Buttons: Pull (Excel → Rhino), Push (Rhino → Excel), Open in Excel (optional: open file path without interop).

- **ExcelLinkRelink**
  - Change file/range and metadata, then re-render.

- **ExcelLinkExportCsv** (optional)
  - Export the current block’s table as CSV.

---

#### Architecture
- **Solution projects**
  - `ExcelLink.Plugin`: Rhino plug-in, commands, and Eto UI
  - `ExcelLink.Core`: ClosedXML-based services (read/write, table normalization, format mapping) — unit-testable

- **NuGet dependencies**
  - `ClosedXML` (preferred) or `EPPlus`
  - `DocumentFormat.OpenXml` if needed

- **Data model**
  - `TableModel → Row → Cell` (+ type, value, alignment, merge info)

- **Renderer**
  - Deterministically converts `TableModel` into Rhino geometry and (re)creates the block definition.

- **Persistence**
  - Store metadata under `InstanceDefinition.UserDictionary["ExcelLink"]` (path, sheet, range, options, version).

---

#### Formatting & Features (MVP)
- Convert Excel column widths to mm (configurable scaling).
- Basic cell alignment (Left/Center/Right, Top/Middle/Bottom).
- Row heights; optional word wrapping (soft line breaks at max width).
- Merged cells: one `TextEntity`; adjust grid accordingly.
- Font from Rhino `DimStyle` or explicitly selectable.

---

#### Samples & Tests
- Include `samples/Example.xlsx` (3 sheets, named range `BOM`).
- Unit tests in `ExcelLink.Core.Tests`:
  - Read simple/merged cells
  - Write roundtrip (xlsx → model → xlsx)
- Manual tests:
  - Insert in model, copy to layout, modify Excel, run Update → both instances refresh
  - Edit in Panel → Push → verify in Excel

---

#### Build & Packaging
- Visual Studio solution, C# 11, reference RhinoCommon (Rhino 8).
- Create Yak-ready `.rhi`/Yak package with manifest.
- Provide README with installation, shortcuts, limitations (no formula support in the Rhino Panel), and support matrix.

---

#### Acceptance Criteria
- Compiles cleanly and loads in Rhino 8.
- `ExcelLinkInsert` creates a block with a visible table; `ExcelLinkUpdate` reflects changes from Excel.
- `ExcelLinkPanel` displays data and supports Push/Pull.
- No Office interop required; `.xlsx` works without Excel installed.
- Large tables (e.g., 100×20) render within reasonable time; UI remains responsive.

---

#### References
- Rhino Developer Docs (RhinoCommon API, Eto UI): `https://www.rhino3d.com`
- Example discussion/plugins around tables in Rhino:
  - McNeel Forum: `https://discourse.mcneel.com`
  - food4rhino: `https://www.food4rhino.com`
- Background on Excel I/O concepts (RhinoScript example): `https://www.rhino3d.com` (context only)
- OLE busy prompt context (we avoid OLE): `https://www.rhino3d.com`

---

#### Deliverables
- Full source (2 projects + tests)
- Compiled plug-in + Yak package
- README with GIF/graphic showing Insert → Update → Push workflow

Prompt für AI Coding Agent
Rolle & Ziel
Du bist ein Senior Rhino/ .NET-Entwickler. Erzeuge ein Rhino 8 (Windows) C#-Plugin namens ExcelLink (RhinoCommon, .NET 7), mit dem ich dynamische Excel-Tabellen in Rhino einfügen, aktualisieren und zurückschreiben kann. Tabellen werden als Block aus Linien/Text gerendert und sind mit einem Excel-Bereich verknüpft (Dateipfad + Sheet + Range/Namensbereich). Keine Office-Interop, nur ClosedXML (oder EPPlus).

Wichtige Hinweise / Best Practices (unbedingt einhalten)

Keine COM/OLE-Interop zu Excel. Verwende ClosedXML (NuGet) zum Lesen/Schreiben von .xlsx. (Grund: Stabilität, keine OLE-Busy-Prompts). 
Stack Overflow

UI mit Eto.Forms (Rhino-konform). Keine WPF-/WinForms-Sonderwege.

Lange Operationen nicht im UI-Thread – RhinoApp.InvokeOnUiThread nur für UI.

Änderungen in einem Undo-Record bündeln.

Block-Definitionen nutzen, um Layout/Model-Instanzen synchron zu halten.

UserDictionary/Archivdaten am Block speichern (Dateipfad, Sheet, Range, Formatierungsoptionen).

Dateiwächter (FileSystemWatcher) plus „Update?“-Prompt.

Kreuzplattform-Tauglichkeit wahren (nur Excel-Teil ist Windows-agnostisch).

Saubere Fehlerbehandlung + Logs in der Rhino-Befehlszeile.

Funktionen (Befehle)

ExcelLinkInsert

Dialog: Excel-Datei (.xlsx), Sheet wählen, Range oder Named Range eingeben, Ziel (Modell oder bestehendes Layout-Blatt), Maßstab (mm), Schriftstil/Height, Optionen (Rasterlinien ja/nein, Textausrichtung, Rahmen).

Lies Zellen via ClosedXML, materialisiere Tabelle: Rechtecke/Linien für Gitter, TextEntity je Zelle, alles in BlockDefinition ExcelLink_<hash>, füge BlockInstance ein.

Speichere Link-Metadaten im Block (Pfad, Sheet, Range, Formatoptionen).

ExcelLinkUpdate

Für gewählte(n) ExcelLink-Block: Excel erneut laden, BlockDefinition rekonstruieren, Instanzen bleiben positioniert.

Option „Alle aktualisieren“.

ExcelLinkPanel (modelloses Panel)

Eto-Grid, das die verknüpfte Tabelle anzeigt und bearbeiten lässt (Basis-Edit: Text/Number, kein Formelsupport).

Buttons: Pull (Excel→Rhino), Push (Rhino→Excel), Open in Excel (optional: nur Dateipfad öffnen, ohne Interop).

ExcelLinkRelink

Datei/Range ändern, Metadaten aktualisieren, dann neu rendern.

ExcelLinkExportCsv (optional)

Aktuellen Block als CSV exportieren.

Architektur

Projekte:

ExcelLink.Plugin (Rhino Plug-in + Commands + UI)

ExcelLink.Core (ClosedXML-basierte Services: Lesen/Schreiben, Tabellennormalisierung, Format-Mapping) – unit-testbar

NuGet: ClosedXML (oder EPPlus), DocumentFormat.OpenXml falls nötig.

Datenmodell: TableModel → Row → Cell (+ Typ, Value, Alignment, MergeInfo).

Renderer: wandelt TableModel in Rhino-Geometrien + erstellt/aktualisiert BlockDefinition deterministisch.

Persistenz: InstanceDefinition.UserDictionary["ExcelLink"] (Pfad, Sheet, Range, Optionen, Version).

Formatierung & Features (MVP)

Spaltenbreiten aus Excel in mm umrechnen (Skalierung konfigurierbar).

Grundlegende Zell-Ausrichtung (Left/Center/Right, Top/Middle/Bottom).

Zeilenhöhen, Auto-Wrap optional (vereinfachen: weiche Zeilenumbrüche bei max Breite).

Zusammengeführte Zellen: als ein TextEntity, Gitter entsprechend anpassen.

Schrift aus Rhino-DimStyle oder explizit wählbar.

Tests & Beispiel

Lege eine samples\Example.xlsx bei (3 Sheets, Named Range „BOM“).

Unit-Tests in ExcelLink.Core.Tests:

Lesen einfacher/zusammengeführter Zellen

Schreib-Roundtrip (xlsx → model → xlsx)

Manuelle Tests:

Insert im Modell, Kopie im Layout, Excel ändern, Update ausführen → beide Instanzen aktualisieren.

Panel-Edit → Push → Excel prüfen.

Build/Packaging

VS-Solution, C# 11, RhinoCommon als Referenz.

Yak-fähiges .rhi/Yak-Package mit manifest.

README mit Installation, Shortcuts, Limitations (kein Formelsupport in Rhino-Panel), Support-Matrix.

Akzeptanzkriterien

Kompiliert fehlerfrei, Plugin lädt in Rhino 8.

ExcelLinkInsert erzeugt Block mit sichtbarer Tabelle; ExcelLinkUpdate spiegelt Excel-Änderungen.

ExcelLinkPanel zeigt Daten und kann „Push“/„Pull“.

Kein Office-Interop erforderlich; .xlsx funktioniert ohne Excel-Installation.

Große Tabellen (z. B. 100×20) rendern innerhalb sinnvoller Zeit; UI bleibt responsiv.

Dokumentation/Links (Referenzen bei Unsicherheit)

Rhino Developer Docs (RhinoCommon API, Samples, Eto-UI). 
www.rhino3d.com
+1

RhinoScript Excel-Lesen (nur als Hintergrund/Beispiel für Excel-I/O-Konzepte). 
www.rhino3d.com
+1

OLE-Busy-Hinweis (nur Kontext, wir meiden OLE). 
www.rhino3d.com

Diskussion/Plugins zum Thema Tabellen in Rhino. 
McNeel Forum
food4rhino.com

Lieferumfang

Vollständiger Sourcecode (2 Projekte + Tests)

Kompilat + Yak-Paket

README mit GIF/Grafik (Insert→Update→Push Workflow)