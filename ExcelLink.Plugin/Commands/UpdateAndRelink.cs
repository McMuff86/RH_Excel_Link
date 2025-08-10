using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Input;
using Rhino.Input.Custom;
using Rhino.Geometry;
using ExcelLink.Core.Rendering;
using ExcelLink.Plugin.Rendering;

namespace ExcelLink.Plugin.Commands;

public static class ExcelLinkUpdater
{
    public static void UpdateDefinition(RhinoDoc doc, InstanceDefinition def, bool extendRangeIfNeeded = false)
    {
        var file = def.GetUserString(BlockMetadataKeys.File);
        var sheet = def.GetUserString(BlockMetadataKeys.Sheet);
        var range = def.GetUserString(BlockMetadataKeys.Range);
        var scaleStr = def.GetUserString(BlockMetadataKeys.Scale);
        var alignStr = def.GetUserString(BlockMetadataKeys.Align);
        var gridStr = def.GetUserString(BlockMetadataKeys.ShowGrid);
        var textMmStr = def.GetUserString(BlockMetadataKeys.TextMm);
        var fontStr = def.GetUserString(BlockMetadataKeys.Font);
        var wrapStr = def.GetUserString(BlockMetadataKeys.Wrap);
        var dimStyleStr = def.GetUserString(BlockMetadataKeys.DimStyle);
        var vAlignStr = def.GetUserString(BlockMetadataKeys.VAlign);
        double scale = 1.0;
        if (!string.IsNullOrWhiteSpace(scaleStr))
            double.TryParse(scaleStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out scale);
        int alignVal = 0;
        if (!string.IsNullOrWhiteSpace(alignStr))
            int.TryParse(alignStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out alignVal);
        bool drawGrid = gridStr == "1";
        double? textMm = null; if (!string.IsNullOrWhiteSpace(textMmStr) && double.TryParse(textMmStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var th)) textMm = th;
        string? fontFamily = string.IsNullOrWhiteSpace(fontStr) ? null : fontStr;
        bool wrap = wrapStr == "1";
        int vAlign = 1; if (!string.IsNullOrWhiteSpace(vAlignStr)) int.TryParse(vAlignStr, out vAlign);

        if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
            return;

        var instances = doc.Objects.GetObjectList(ObjectType.InstanceReference)
            .OfType<InstanceObject>()
            .Where(io => io.InstanceDefinition?.Index == def.Index)
            .ToArray();
        var xforms = instances.Select(i => i.InstanceXform).ToArray();
        foreach (var inst in instances)
            doc.Objects.Delete(inst, true);

        var reader = new ExcelLink.Core.Rendering.ExcelClosedXmlReader();
        var table = reader.ReadTable(file, sheet, range);

        // If used range extends beyond stored range and user opted-in, update range
        if (extendRangeIfNeeded)
        {
            var used = ExcelLink.Core.Services.ExcelInspector.GetUsedRangeA1(file, sheet);
            if (!string.IsNullOrWhiteSpace(used) && !string.Equals(used, range, System.StringComparison.OrdinalIgnoreCase))
            {
                range = used;
                def.SetUserString(BlockMetadataKeys.Range, range);
                table = reader.ReadTable(file, sheet, range);
            }
        }
        ExcelLink.Plugin.Rendering.RhinoTableRenderer.RenderTableAsBlock(doc, table, def.Name, Rhino.Geometry.Point3d.Origin, scale, true,
            (ExcelLink.Plugin.Rendering.RhinoTableRenderer.HorizontalOverride)alignVal, vAlign,
            false, out var newDefIndex, out _, out var newInstanceIds, xforms,
            drawGrid, textMm, fontFamily, dimStyleStr, wrap);

        var newDef = doc.InstanceDefinitions[newDefIndex];
        // Reapply metadata to new definition (keep keys stable including DimStyle)
        newDef.SetUserString(BlockMetadataKeys.File, file);
        newDef.SetUserString(BlockMetadataKeys.Sheet, sheet);
        newDef.SetUserString(BlockMetadataKeys.Range, range);
        newDef.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
        newDef.SetUserString(BlockMetadataKeys.Align, alignVal.ToString(System.Globalization.CultureInfo.InvariantCulture));
        newDef.SetUserString(BlockMetadataKeys.ShowGrid, drawGrid ? "1" : "0");
        if (textMm.HasValue) newDef.SetUserString(BlockMetadataKeys.TextMm, textMm.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        else newDef.SetUserString(BlockMetadataKeys.TextMm, string.Empty);
        newDef.SetUserString(BlockMetadataKeys.Font, fontFamily ?? string.Empty);
        newDef.SetUserString(BlockMetadataKeys.Wrap, wrap ? "1" : "0");
        newDef.SetUserString(BlockMetadataKeys.VAlign, vAlign.ToString(System.Globalization.CultureInfo.InvariantCulture));
        newDef.SetUserString(BlockMetadataKeys.DimStyle, dimStyleStr ?? string.Empty);
        // Update watcher mapping to the new index
        ExcelLink.Plugin.Services.ExcelLinkWatcherService.Instance.ReplaceDefinitionIndex(file, def.Index, newDefIndex);

        // Reapply metadata to new instances explicitly and ensure DimStyle key persists
        if (newInstanceIds != null && newInstanceIds.Count > 0)
        {
            foreach (var id in newInstanceIds)
            {
                var inst = doc.Objects.FindId(id) as InstanceObject;
                if (inst == null) continue;
                var a = inst.Attributes.Duplicate();
                a.SetUserString(BlockMetadataKeys.File, file);
                a.SetUserString(BlockMetadataKeys.Sheet, sheet);
                a.SetUserString(BlockMetadataKeys.Range, range);
                a.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
                a.SetUserString(BlockMetadataKeys.Align, alignVal.ToString(System.Globalization.CultureInfo.InvariantCulture));
                a.SetUserString(BlockMetadataKeys.ShowGrid, drawGrid ? "1" : "0");
                if (textMm.HasValue) a.SetUserString(BlockMetadataKeys.TextMm, textMm.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                else a.SetUserString(BlockMetadataKeys.TextMm, string.Empty);
                a.SetUserString(BlockMetadataKeys.Font, fontFamily ?? string.Empty);
                a.SetUserString(BlockMetadataKeys.Wrap, wrap ? "1" : "0");
                a.SetUserString(BlockMetadataKeys.VAlign, vAlign.ToString(System.Globalization.CultureInfo.InvariantCulture));
                a.SetUserString(BlockMetadataKeys.DimStyle, dimStyleStr ?? string.Empty);
                doc.Objects.ModifyAttributes(inst, a, true);
            }
        }
    }
}

internal static class BlockMetadataKeys
{
    public const string Prefix = "ExcelLink_";
    public const string File = Prefix + "File";
    public const string Sheet = Prefix + "Sheet";
    public const string Range = Prefix + "Range";
    public const string Scale = Prefix + "Scale";
    public const string Align = Prefix + "Align";
    public const string VAlign = Prefix + "VAlign"; // 0=Top,1=Middle,2=Bottom
    public const string ShowGrid = Prefix + "ShowGrid";
    public const string Grid = Prefix + "Grid";
    public const string TextMm = Prefix + "TextMm";
    public const string Font = Prefix + "Font";
    public const string Wrap = Prefix + "Wrap";
    public const string DimStyle = Prefix + "DimStyle";
}

[Guid("f6b5c7f5-9c1a-4d5e-9d51-1f0f4d3f0b90")]
public sealed class ExcelLinkUpdateCommand : Command
{
    public override string EnglishName => "ExcelLinkUpdate";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var go = new GetObject();
        go.SetCommandPrompt("Select ExcelLink block instance(s) to update");
        go.GeometryFilter = ObjectType.InstanceReference;
        go.GetMultiple(1, 0);
        if (go.CommandResult() != Result.Success)
            return go.CommandResult();

        foreach (var objRef in go.Objects())
        {
            if (objRef.Object() is not InstanceObject iref)
                continue;
            var def = iref.InstanceDefinition;
            if (def == null)
                continue;

            var defName = def.Name;
            var defIndex = def.Index;

            // Keep metadata before redefining
            var file = def.GetUserString(BlockMetadataKeys.File);
            var sheet = def.GetUserString(BlockMetadataKeys.Sheet);
            var range = def.GetUserString(BlockMetadataKeys.Range);
            var scaleStr = def.GetUserString(BlockMetadataKeys.Scale);
            var alignStr = def.GetUserString(BlockMetadataKeys.Align);
            var vAlignStr = def.GetUserString(BlockMetadataKeys.VAlign);
            double scale = 1.0;
            if (!string.IsNullOrWhiteSpace(scaleStr))
                double.TryParse(scaleStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out scale);
            int alignVal = 0;
            if (!string.IsNullOrWhiteSpace(alignStr))
                int.TryParse(alignStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out alignVal);
            int vAlign = 1;
            if (!string.IsNullOrWhiteSpace(vAlignStr))
                int.TryParse(vAlignStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out vAlign);

            if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
                continue;

            try
            {
                var reader = new ExcelClosedXmlReader();
                var table = reader.ReadTable(file, sheet, range);
                // collect transforms of all instances of this definition
                var instanceObjects = doc.Objects.GetObjectList(ObjectType.InstanceReference)
                    .OfType<InstanceObject>()
                    .Where(io => io.InstanceDefinition?.Index == defIndex)
                    .ToArray();
                var xforms = instanceObjects.Select(i => i.InstanceXform).ToArray();

                // Delete existing instances before redefining to avoid dangling instances referencing old def
                foreach (var inst in instanceObjects)
                    doc.Objects.Delete(inst, true);

                RhinoTableRenderer.RenderTableAsBlock(doc, table, defName, iref.InsertionPoint,
                    scale, true, (RhinoTableRenderer.HorizontalOverride)alignVal, vAlign,
                    false, out var newDefIndex, out _, out var newIds, xforms,
                    def.GetUserString(BlockMetadataKeys.ShowGrid) == "1",
                    double.TryParse(def.GetUserString(BlockMetadataKeys.TextMm), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var th) ? th : (double?)null,
                    def.GetUserString(BlockMetadataKeys.Font),
                    def.GetUserString(BlockMetadataKeys.DimStyle),
                    def.GetUserString(BlockMetadataKeys.Wrap) == "1");

                // Reapply metadata to the new definition
                var newDef = doc.InstanceDefinitions[newDefIndex];
                if (newDef != null)
                {
                    newDef.SetUserString(BlockMetadataKeys.File, file ?? string.Empty);
                    newDef.SetUserString(BlockMetadataKeys.Sheet, sheet ?? string.Empty);
                    newDef.SetUserString(BlockMetadataKeys.Range, range ?? string.Empty);
                    newDef.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    newDef.SetUserString(BlockMetadataKeys.Align, alignVal.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    newDef.SetUserString(BlockMetadataKeys.ShowGrid, def.GetUserString(BlockMetadataKeys.ShowGrid));
                    newDef.SetUserString(BlockMetadataKeys.TextMm, def.GetUserString(BlockMetadataKeys.TextMm));
                    newDef.SetUserString(BlockMetadataKeys.Font, def.GetUserString(BlockMetadataKeys.Font));
                    newDef.SetUserString(BlockMetadataKeys.Wrap, def.GetUserString(BlockMetadataKeys.Wrap));
                    newDef.SetUserString(BlockMetadataKeys.DimStyle, def.GetUserString(BlockMetadataKeys.DimStyle));
                }
                // Update watcher mapping
                ExcelLink.Plugin.Services.ExcelLinkWatcherService.Instance.ReplaceDefinitionIndex(file ?? string.Empty, defIndex, newDefIndex);

                // Reapply metadata to all instances of the new definition (prefer newly created ids)
                var reinserts = (newIds != null && newIds.Count > 0)
                    ? newIds.Select(id => doc.Objects.FindId(id)).OfType<InstanceObject>().ToArray()
                    : doc.Objects.GetObjectList(ObjectType.InstanceReference)
                        .OfType<InstanceObject>()
                        .Where(io => io.InstanceDefinition?.Index == newDefIndex)
                        .ToArray();
                foreach (var inst in reinserts)
                {
                    var a = inst.Attributes.Duplicate();
                    a.SetUserString(BlockMetadataKeys.File, file ?? string.Empty);
                    a.SetUserString(BlockMetadataKeys.Sheet, sheet ?? string.Empty);
                    a.SetUserString(BlockMetadataKeys.Range, range ?? string.Empty);
                    a.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    a.SetUserString(BlockMetadataKeys.Align, alignVal.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    doc.Objects.ModifyAttributes(inst, a, true);
                }

                RhinoApp.WriteLine($"Updated '{defName}'.");
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Update failed: {ex.Message}");
            }
        }

        doc.Views.Redraw();
        return Result.Success;
    }
}

[Guid("4c9cbb7d-7f9a-4d88-9d2a-5a7c9bb3a7cb")]
public sealed class ExcelLinkRelinkCommand : Command
{
    public override string EnglishName => "ExcelLinkRelink";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var go = new GetObject();
        go.SetCommandPrompt("Select ExcelLink block instance to relink");
        go.GeometryFilter = ObjectType.InstanceReference;
        go.Get();
        if (go.CommandResult() != Result.Success)
            return go.CommandResult();

        if (go.ObjectCount != 1)
            return Result.Cancel;

        var iref = go.Object(0).Object() as InstanceObject;
        if (iref == null) return Result.Failure;
        var def = iref.InstanceDefinition;
        if (def == null) return Result.Failure;

        var idlg = new UI.InsertDialog();
        idlg.Title = "Relink Excel";
        var res = idlg.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow);
        if (idlg.Result == null) return Result.Cancel;
        var file = idlg.Result.FilePath;
        var sheet = idlg.Result.Sheet;
        var range = idlg.Result.RangeOrName;

        if (!File.Exists(file))
        {
            RhinoApp.WriteLine($"File not found: {file}");
            return Result.Failure;
        }

        // Update metadata
        def.SetUserString(BlockMetadataKeys.File, file);
        def.SetUserString(BlockMetadataKeys.Sheet, sheet);
        def.SetUserString(BlockMetadataKeys.Range, range);

        RhinoApp.WriteLine($"Relinked '{def.Name}'. Run ExcelLinkUpdate to refresh.");
        return Result.Success;
    }
}


