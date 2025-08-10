using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Input.Custom;

namespace ExcelLink.Plugin.Commands;

[Guid("3a7d3b3a-6b5f-4b2f-8f8e-7ea1b0c4a4f2")]
public sealed class ExcelLinkApplyAttributesCommand : Command
{
    public override string EnglishName => "ExcelLinkApplyAttributes";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var go = new GetObject();
        go.SetCommandPrompt("Select ExcelLink block instance(s) to apply attributes");
        go.GeometryFilter = ObjectType.InstanceReference;
        go.GetMultiple(1, 0);
        if (go.CommandResult() != Result.Success)
            return go.CommandResult();

        foreach (var objRef in go.Objects())
        {
            if (objRef.Object() is not InstanceObject inst)
                continue;
            var def = inst.InstanceDefinition;
            if (def == null)
                continue;

            // Apply user strings from instance (if present) to definition
            var keys = new[]
            {
                BlockMetadataKeys.File,
                BlockMetadataKeys.Sheet,
                BlockMetadataKeys.Range,
                BlockMetadataKeys.Scale,
                BlockMetadataKeys.Align,
                BlockMetadataKeys.ShowGrid,
                BlockMetadataKeys.TextMm,
                BlockMetadataKeys.Font,
                BlockMetadataKeys.Wrap,
                BlockMetadataKeys.DimStyle
            };

            foreach (var key in keys)
            {
                var val = inst.Attributes.GetUserString(key);
                if (!string.IsNullOrEmpty(val))
                    def.SetUserString(key, val);
            }

            try
            {
                ExcelLinkUpdater.UpdateDefinition(doc, def, extendRangeIfNeeded: false);
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"ApplyAttributes failed: {ex.Message}");
            }
        }

        doc.Views.Redraw();
        return Result.Success;
    }
}


