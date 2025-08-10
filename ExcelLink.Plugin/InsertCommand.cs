using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rhino;
using Rhino.Commands;
using Rhino.Input;
using Rhino.Input.Custom;
using ExcelLink.Core.Rendering;
using ExcelLink.Plugin.Rendering;
using ExcelLink.Plugin.Commands;

namespace ExcelLink.Plugin;

[Guid("0a6c7a84-a82d-4a19-9c83-3f52f6c5c0d7")]
public sealed class ExcelLinkInsertCommand : Command
{
    public override string EnglishName => "ExcelLinkInsert";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var gs = new GetString();
        gs.SetCommandPrompt("Enter Excel spec: <filePath>|<sheet>|<rangeOrNamedRange>");
        gs.AcceptNothing(false);
        var res = gs.Get();
        if (res != GetResult.String)
            return gs.CommandResult();

        var spec = (gs.StringResult() ?? string.Empty).Trim();
        if (!TryParseSpec(spec, out var filePath, out var sheet, out var range))
        {
            RhinoApp.WriteLine("Invalid format. Example: C:\\Users\\me\\file.xlsx|Sheet1|A1:D20");
            return Result.Failure;
        }

        if (!File.Exists(filePath))
        {
            RhinoApp.WriteLine($"File not found: {filePath}");
            return Result.Failure;
        }

        try
        {
            var reader = new ExcelClosedXmlReader();
            var table = reader.ReadTable(filePath, sheet, range);
            RhinoApp.WriteLine($"Loaded {table.Rows.Count} rows × {table.Columns.Count} columns from '{sheet}:{range}'.");

            // Ask for insert point
            var gp = new GetPoint();
            gp.SetCommandPrompt("Pick insert point");
            if (gp.Get() != GetResult.Point)
                return gp.CommandResult();
            var insertPoint = gp.Point();

            // Ask for scale factor (optional)
            double scale = 1.0;
            var gsScale = new GetString();
            gsScale.SetCommandPrompt("Scale (1 = 1:1 in mm to model units). Press Enter for 1");
            gsScale.AcceptNothing(true);
            var scaleRes = gsScale.Get();
            if (scaleRes == GetResult.String && double.TryParse(gsScale.StringResult(), out var parsed))
                scale = Math.Max(0.001, parsed);

            // Ask for horizontal alignment override
            var gsAlign = new GetOption();
            gsAlign.SetCommandPrompt("Horizontal alignment (Enter = From Excel)");
            var optLeft = gsAlign.AddOption("Left");
            var optCenter = gsAlign.AddOption("Center");
            var optRight = gsAlign.AddOption("Right");
            var optExcel = gsAlign.AddOption("Excel");
            RhinoTableRenderer.HorizontalOverride horiz = RhinoTableRenderer.HorizontalOverride.UseExcel;
            var resAlign = gsAlign.Get();
            if (resAlign == GetResult.Option)
            {
                if (gsAlign.OptionIndex() == optLeft) horiz = RhinoTableRenderer.HorizontalOverride.Left;
                else if (gsAlign.OptionIndex() == optCenter) horiz = RhinoTableRenderer.HorizontalOverride.Center;
                else if (gsAlign.OptionIndex() == optRight) horiz = RhinoTableRenderer.HorizontalOverride.Right;
                else horiz = RhinoTableRenderer.HorizontalOverride.UseExcel;
            }

            var blockName = $"ExcelLink_{Guid.NewGuid():N}";
            var rr = RhinoTableRenderer.RenderTableAsBlock(
                doc,
                table,
                blockName,
                insertPoint,
                scaleMultiplier: scale,
                topRowAtTop: true,
                horizontal: horiz,
                insertInstance: true,
                out var defIndex,
                out var instanceId,
                reinsertTransforms: null);
            if (rr == Result.Success)
            {
                RhinoApp.WriteLine($"Inserted block '{blockName}' at {insertPoint} (scale {scale}). DefIndex={defIndex}");

                // Persist link metadata on instance definition
                var def = doc.InstanceDefinitions.Find(blockName);
                if (def != null)
                {
                    def.SetUserString(BlockMetadataKeys.File, filePath);
                    def.SetUserString(BlockMetadataKeys.Sheet, sheet);
                    def.SetUserString(BlockMetadataKeys.Range, range);
                    def.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    def.SetUserString(BlockMetadataKeys.Align, ((int)horiz).ToString(System.Globalization.CultureInfo.InvariantCulture));
                    // Also on the inserted instance (helps selection scenarios)
                    if (instanceId != Guid.Empty)
                    {
                        var obj = doc.Objects.FindId(instanceId);
                        if (obj != null)
                        {
                            var attrs = obj.Attributes.Duplicate();
                            attrs.SetUserString(BlockMetadataKeys.File, filePath);
                            attrs.SetUserString(BlockMetadataKeys.Sheet, sheet);
                            attrs.SetUserString(BlockMetadataKeys.Range, range);
                            attrs.SetUserString(BlockMetadataKeys.Scale, scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
                            attrs.SetUserString(BlockMetadataKeys.Align, ((int)horiz).ToString(System.Globalization.CultureInfo.InvariantCulture));
                            doc.Objects.ModifyAttributes(obj, attrs, true);
                        }
                    }
                }
            }
            return rr;
        }
        catch (Exception ex)
        {
            RhinoApp.WriteLine($"Error reading Excel: {ex.Message}");
            return Result.Failure;
        }
    }

    public static bool TryParseSpec(string spec, out string filePath, out string sheet, out string rangeOrName)
    {
        filePath = string.Empty;
        sheet = string.Empty;
        rangeOrName = string.Empty;
        if (string.IsNullOrWhiteSpace(spec)) return false;

        var parts = spec.Split('|');
        if (parts.Length < 3) return false;

        filePath = parts[0].Trim().Trim('"');
        sheet = parts[1].Trim();
        rangeOrName = string.Join("|", parts.Skip(2)).Trim();
        return !(string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(rangeOrName));
    }
}


