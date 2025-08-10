### ExcelLink Architecture Overview

This document describes the high-level architecture for the ExcelLink Rhino plug-in.

---

#### Solution Structure
- `ExcelLink.Plugin`
  - Rhino plug-in entry point
  - Commands: `ExcelLinkInsert`, `ExcelLinkUpdate`, `ExcelLinkRelink`, `ExcelLinkExportCsv`
  - UI: Eto.Forms views and a modeless `ExcelLinkPanel`
  - Rendering: orchestrates geometry creation and block updates

- `ExcelLink.Core`
  - Excel I/O services using ClosedXML
  - Table normalization and format mapping
  - Data models (`TableModel`, `Row`, `Cell`)
  - Mappers between Excel styles and Rhino text/geometry options
  - Unit tests in `ExcelLink.Core.Tests`

---

#### Data Model
- `TableModel`
  - `Rows: List<Row>`
  - `Columns: List<ColumnSpec>` (optional width info)
  - `Styles: TableStyle` (defaults)

- `Row`
  - `Cells: List<Cell>`
  - `HeightMm: double`

- `Cell`
  - `Value: string | number`
  - `DataType: enum { Text, Number, Date, Bool }`
  - `HorizontalAlignment: enum { Left, Center, Right }`
  - `VerticalAlignment: enum { Top, Middle, Bottom }`
  - `Merge: MergeInfo?` (row/column span)
  - `Style: CellStyle` (font, size, bold, italic, wrap, borders)

---

#### Core Services
- `IExcelReader`
  - Reads an Excel range (file, sheet, range or named range) into `TableModel`

- `IExcelWriter`
  - Writes `TableModel` back to Excel at the specified range

- `ITableNormalizer`
  - Normalizes merged cells, alignment, widths/heights, wraps to fit

- `ITableRenderer`
  - Converts `TableModel` to Rhino geometry (grid lines, text entities)
  - Produces/updates a deterministic block definition

- `ILinkMetadataStore`
  - Persists link metadata on `InstanceDefinition.UserDictionary["ExcelLink"]`

---

#### Rendering Pipeline
1. Read Excel range → `TableModel`
2. Normalize model (sizes, merges, wrap)
3. Map to Rhino text styles and units
4. Build geometry (lines/rectangles + text entities)
5. Create/update block definition `ExcelLink_<hash>`
6. Insert/refresh block instances in model/layout

---

#### Concurrency & Responsiveness
- Perform file I/O and heavy processing off the UI thread
- Use `RhinoApp.InvokeOnUiThread` only to manipulate UI or Rhino document state
- Group changes under a single Undo record per update

---

#### File Watching & Updates
- `FileSystemWatcher` observes source `.xlsx`
- On change: prompt user to update affected blocks
- Provide command to update all linked instances

---

#### Error Handling & Logging
- Wrap Excel I/O and rendering with try/catch boundaries
- Log concise messages to Rhino command line
- Provide actionable messages in dialogs/panel

---

#### Packaging & Deployment
- Build with .NET 7, C# 11, reference RhinoCommon (Rhino 8)
- Create Yak package with manifest and `.rhi`
- Include `samples/Example.xlsx` and unit tests


