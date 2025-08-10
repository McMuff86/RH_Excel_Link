using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelLink.Core.Services;

public static class ExcelInspector
{
    public static List<string> ListSheetNames(string filePath)
    {
        using var wb = new XLWorkbook(filePath);
        var list = new List<string>();
        foreach (var ws in wb.Worksheets)
            list.Add(ws.Name);
        return list;
    }

    public static string? GetUsedRangeA1(string filePath, string sheetName)
    {
        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet(sheetName);
        if (ws == null) return null;
        var used = ws.RangeUsed();
        if (used == null) return null;
        var first = used.RangeAddress.FirstAddress;
        var last = used.RangeAddress.LastAddress;
        return $"{first.ToStringRelative()}:{last.ToStringRelative()}";
    }
}


