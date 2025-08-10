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

internal static class BlockMetadataKeys
{
    public const string Prefix = "ExcelLink_";
    public const string File = Prefix + "File";
    public const string Sheet = Prefix + "Sheet";
    public const string Range = Prefix + "Range";
    public const string Scale = Prefix + "Scale";
    public const string Align = Prefix + "Align";
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
            double scale = 1.0;
            if (!string.IsNullOrWhiteSpace(scaleStr))
                double.TryParse(scaleStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out scale);
            int alignVal = 0;
            if (!string.IsNullOrWhiteSpace(alignStr))
                int.TryParse(alignStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out alignVal);

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

                RhinoTableRenderer.RenderTableAsBlock(doc, table, defName, iref.InsertionPoint, scale, true,
                    (RhinoTableRenderer.HorizontalOverride)alignVal, insertInstance: false, out var newDefIndex, out _, reinsertTransforms: xforms);

                // Reapply metadata to the new definition
                var newDef = doc.InstanceDefinitions[newDefIndex];
                if (newDef != null)
                {
                    newDef.SetUserString(BlockMetadataKeys.File, file ?? string.Empty);
                    newDef.SetUserString(BlockMetadataKeys.Sheet, sheet ?? string.Empty);
                    newDef.SetUserString(BlockMetadataKeys.Range, range ?? string.Empty);
                    newDef.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    newDef.SetUserString(BlockMetadataKeys.Align, alignVal.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }

                // Reapply metadata to all instances of the new definition
                var reinserts = doc.Objects.GetObjectList(ObjectType.InstanceReference)
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

        var gspec = new GetString();
        gspec.SetCommandPrompt("New Excel spec: <filePath>|<sheet>|<rangeOrNamedRange>");
        if (gspec.Get() != GetResult.String) return gspec.CommandResult();
        var spec = (gspec.StringResult() ?? string.Empty).Trim();
        if (!ExcelLink.Plugin.ExcelLinkInsertCommand.TryParseSpec(spec, out var file, out var sheet, out var range))
            return Result.Failure;

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


