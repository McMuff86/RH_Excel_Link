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
        // UI dialog for input
        var dlg = new UI.InsertDialog();
        var dlgRes = dlg.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow);
        if (dlg.Result == null)
            return Result.Cancel;
        var filePath = dlg.Result.FilePath;
        var sheet = dlg.Result.Sheet;
        var range = dlg.Result.RangeOrName;

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

            // Pick insert point
            var gp = new GetPoint();
            gp.SetCommandPrompt("Pick insert point");
            if (gp.Get() != GetResult.Point)
                return gp.CommandResult();
            var insertPoint = gp.Point();

            // Scale & alignment from dialog
            double scale = Math.Max(0.001, dlg.Result.Scale);
            RhinoTableRenderer.HorizontalOverride horiz = dlg.Result.Alignment switch
            {
                "Left" => RhinoTableRenderer.HorizontalOverride.Left,
                "Center" => RhinoTableRenderer.HorizontalOverride.Center,
                "Right" => RhinoTableRenderer.HorizontalOverride.Right,
                _ => RhinoTableRenderer.HorizontalOverride.UseExcel
            };

            var blockName = $"ExcelLink_{Guid.NewGuid():N}";
            var rr = RhinoTableRenderer.RenderTableAsBlock(
                doc,
                table,
                blockName,
                insertPoint,
                scale,
                true,
                horiz,
                1,
                true,
                out var defIndex,
                out var instanceId,
                out var _,
                null,
                dlg.Result.DrawGrid,
                dlg.Result.TextHeightMm,
                dlg.Result.FontFamily,
                dlg.Result.DimStyleName,
                dlg.Result.Wrap);
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
                    def.SetUserString(BlockMetadataKeys.ShowGrid, dlg.Result.DrawGrid ? "1" : "0");
                    if (dlg.Result.TextHeightMm.HasValue)
                        def.SetUserString(BlockMetadataKeys.TextMm, dlg.Result.TextHeightMm.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    else
                        def.SetUserString(BlockMetadataKeys.TextMm, string.Empty);
                    if (!string.IsNullOrWhiteSpace(dlg.Result.FontFamily))
                        def.SetUserString(BlockMetadataKeys.Font, dlg.Result.FontFamily);
                    else
                        def.SetUserString(BlockMetadataKeys.Font, string.Empty);
                    def.SetUserString(BlockMetadataKeys.DimStyle, dlg.Result.DimStyleName ?? string.Empty);
                    def.SetUserString(BlockMetadataKeys.Wrap, dlg.Result.Wrap ? "1" : "0");
                    def.SetUserString(BlockMetadataKeys.VAlign, "1");
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
                            attrs.SetUserString(BlockMetadataKeys.ShowGrid, dlg.Result.DrawGrid ? "1" : "0");
                            if (dlg.Result.TextHeightMm.HasValue)
                                attrs.SetUserString(BlockMetadataKeys.TextMm, dlg.Result.TextHeightMm.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                            else
                                attrs.SetUserString(BlockMetadataKeys.TextMm, string.Empty);
                            if (!string.IsNullOrWhiteSpace(dlg.Result.FontFamily))
                                attrs.SetUserString(BlockMetadataKeys.Font, dlg.Result.FontFamily);
                            else
                                attrs.SetUserString(BlockMetadataKeys.Font, string.Empty);
                            attrs.SetUserString(BlockMetadataKeys.DimStyle, dlg.Result.DimStyleName ?? string.Empty);
                            attrs.SetUserString(BlockMetadataKeys.Wrap, dlg.Result.Wrap ? "1" : "0");
                            attrs.SetUserString(BlockMetadataKeys.VAlign, "1");
                            doc.Objects.ModifyAttributes(obj, attrs, true);
                        }
                    }
                    // Register watcher
                    ExcelLink.Plugin.Services.ExcelLinkWatcherService.Instance.RegisterDefinition(doc, def.Index, filePath);
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


