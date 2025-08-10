using ClosedXML.Excel;
using ExcelLink.Core.Abstractions;
using ExcelLink.Core.Models;
using System.Linq;

namespace ExcelLink.Core.Rendering;

public sealed class ExcelClosedXmlReader : IExcelReader
{
    public TableModel ReadTable(string filePath, string sheetName, string rangeOrNamedRange)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheet(sheetName);

        IXLRange range;
        try
        {
            range = worksheet.Range(rangeOrNamedRange);
        }
        catch
        {
            var named = workbook.NamedRange(rangeOrNamedRange)
                       ?? worksheet.NamedRange(rangeOrNamedRange);
            if (named == null)
                throw new ArgumentException($"Range or named range '{rangeOrNamedRange}' not found.");

            var first = named.Ranges.First();
            range = first.RangeUsed() ?? first;
        }

        var model = new TableModel();

        var firstColumnNumber = range.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumnNumber = range.RangeAddress.LastAddress.ColumnNumber;
        var firstRowNumber = range.RangeAddress.FirstAddress.RowNumber;
        var lastRowNumber = range.RangeAddress.LastAddress.RowNumber;

        // Column widths from worksheet columns
        for (var c = firstColumnNumber; c <= lastColumnNumber; c++)
        {
            var widthChars = worksheet.Column(c).Width;
            model.Columns.Add(new ColumnSpec
            {
                WidthMillimeters = ConvertColumnWidthToMillimeters(widthChars)
            });
        }

        for (var r = firstRowNumber; r <= lastRowNumber; r++)
        {
            var mRow = new TableRow
            {
                HeightMillimeters = ConvertRowHeightToMillimeters(worksheet.Row(r).Height)
            };

            for (var c = firstColumnNumber; c <= lastColumnNumber; c++)
            {
                var cell = worksheet.Cell(r, c);
                var value = cell.Value;
                var type = InferDataType(cell);
                var alignH = cell.Style.Alignment.Horizontal switch
                {
                    XLAlignmentHorizontalValues.Center => HorizontalAlignment.Center,
                    XLAlignmentHorizontalValues.Right => HorizontalAlignment.Right,
                    _ => HorizontalAlignment.Left
                };
                var alignV = cell.Style.Alignment.Vertical switch
                {
                    XLAlignmentVerticalValues.Top => VerticalAlignment.Top,
                    XLAlignmentVerticalValues.Center => VerticalAlignment.Middle,
                    XLAlignmentVerticalValues.Bottom => VerticalAlignment.Bottom,
                    _ => VerticalAlignment.Middle
                };

                var mCell = new TableCell
                {
                    Value = value,
                    DataType = type,
                    HorizontalAlignment = alignH,
                    VerticalAlignment = alignV,
                    Style = new CellStyle
                    {
                        Bold = cell.Style.Font.Bold,
                        Italic = cell.Style.Font.Italic,
                        WrapText = cell.Style.Alignment.WrapText,
                        FontSizePoints = cell.Style.Font.FontSize
                    }
                };

                if (cell.MergedRange() is { } merged)
                {
                    mCell.Merge = new MergeInfo
                    {
                        RowSpan = merged.RowCount(),
                        ColumnSpan = merged.ColumnCount()
                    };
                }

                mRow.Cells.Add(mCell);
            }

            model.Rows.Add(mRow);
        }

        return model;
    }

    private static CellDataType InferDataType(IXLCell cell)
    {
        return cell.DataType switch
        {
            XLDataType.Number => CellDataType.Number,
            XLDataType.DateTime => CellDataType.Date,
            XLDataType.Boolean => CellDataType.Bool,
            _ => CellDataType.Text
        };
    }

    private static double ConvertColumnWidthToMillimeters(double excelWidthChars)
    {
        // Simple heuristic: Excel column width ~ character width; translate to mm via a fixed factor.
        // 1 character ≈ 2.2 mm for default font/size; this is later refined by normalization.
        return excelWidthChars * 2.2;
    }

    private static double ConvertRowHeightToMillimeters(double excelHeightPoints)
    {
        // 1 point = 1/72 inch; 1 inch = 25.4 mm
        return excelHeightPoints * 25.4 / 72.0;
    }
}


