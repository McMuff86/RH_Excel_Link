using System.Collections.Generic;

namespace ExcelLink.Core.Models;

public sealed class TableModel
{
    public List<TableRow> Rows { get; } = new();
    public List<ColumnSpec> Columns { get; } = new();
    public TableStyle Style { get; set; } = new();
}

public sealed class TableRow
{
    public List<TableCell> Cells { get; } = new();
    public double HeightMillimeters { get; set; }
}

public sealed class ColumnSpec
{
    public double WidthMillimeters { get; set; }
}

public sealed class TableCell
{
    public object? Value { get; set; }
    public CellDataType DataType { get; set; }
    public HorizontalAlignment HorizontalAlignment { get; set; }
    public VerticalAlignment VerticalAlignment { get; set; }
    public MergeInfo? Merge { get; set; }
    public CellStyle Style { get; set; } = new();
}

public enum CellDataType
{
    Text,
    Number,
    Date,
    Bool
}

public enum HorizontalAlignment
{
    Left,
    Center,
    Right
}

public enum VerticalAlignment
{
    Top,
    Middle,
    Bottom
}

public sealed class MergeInfo
{
    public int RowSpan { get; set; } = 1;
    public int ColumnSpan { get; set; } = 1;
}

public sealed class TableStyle
{
    public double DefaultFontSizePoints { get; set; } = 10;
    public string FontFamily { get; set; } = "Arial";
}

public sealed class CellStyle
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool WrapText { get; set; }
    public double? FontSizePoints { get; set; }
}


