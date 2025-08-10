using ExcelLink.Core.Abstractions;
using ExcelLink.Core.Models;

namespace ExcelLink.Core.Rendering;

public sealed class TableNormalizer : ITableNormalizer
{
    public TableModel Normalize(TableModel input)
    {
        // For now, return the model as-is. Future work:
        // - Resolve merged cells layout
        // - Ensure minimum sizes and consistent units
        // - Apply wrapping based on column widths
        return input;
    }
}


