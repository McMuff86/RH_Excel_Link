using ExcelLink.Core.Models;
using ExcelLink.Core.Rendering;
using Xunit;

namespace ExcelLink.Core.Tests;

public class TableNormalizerTests
{
    [Fact]
    public void Normalize_ReturnsSameModel_InstancePreservedForNow()
    {
        var model = new TableModel();
        model.Rows.Add(new TableRow { HeightMillimeters = 5 });
        model.Columns.Add(new ColumnSpec { WidthMillimeters = 10 });
        model.Rows[0].Cells.Add(new TableCell { Value = "A", DataType = CellDataType.Text });

        var normalizer = new TableNormalizer();
        var normalized = normalizer.Normalize(model);

        Assert.Same(model, normalized);
    }
}


