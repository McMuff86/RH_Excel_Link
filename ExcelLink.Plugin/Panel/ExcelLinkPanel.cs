using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Eto.Drawing;
using Eto.Forms;
using Rhino;
using Rhino.DocObjects;

namespace ExcelLink.Plugin.UI;

[Guid("6ddcb0c5-5fb1-4b9a-acdf-9b7e7a6d3a03")]
public sealed class ExcelLinkPanel : Eto.Forms.Panel
{
    public ExcelLinkPanel()
    {
        var filePathText = new TextBox { PlaceholderText = "Excel file path" };
        var sheetText = new TextBox { PlaceholderText = "Sheet" };
        var rangeText = new TextBox { PlaceholderText = "Range or Named Range" };
        var scaleBox = new TextBox { PlaceholderText = "Scale (mm→model)" };
        var alignDrop = new DropDown { DataStore = new[] { "Excel", "Left", "Center", "Right" }, SelectedIndex = 0 };
        var vAlignDrop = new DropDown { DataStore = new[] { "Top", "Middle", "Bottom" }, SelectedIndex = 1 };
        var gridCheck = new CheckBox { Text = "Grid", Checked = true };
        var wrapCheck = new CheckBox { Text = "Wrap", Checked = false };
        var textMm = new TextBox { PlaceholderText = "Text (mm)" };
        var fontBox = new TextBox { PlaceholderText = "Font family" };
        var dimStyleDrop = new DropDown { };
        try
        {
            var doc = Rhino.RhinoDoc.ActiveDoc;
            if (doc != null)
            {
                dimStyleDrop.DataStore = doc.DimStyles.Select(ds => ds.Name).ToList();
            }
        }
        catch { }

        var pullButton = new Button { Text = "Pull (Excel → Rhino)" };
        var updateButton = new Button { Text = "Update Selected" };
        var applyButton = new Button { Text = "Apply Attributes" };

        Content = new StackLayout
        {
            Padding = 10,
            Spacing = 6,
            Items =
            {
                new Label { Text = "ExcelLink Panel", Font = new Eto.Drawing.Font(SystemFont.Bold, 12) },
                filePathText,
                sheetText,
                rangeText,
                new Label { Text = "Options" },
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { new Label{ Text="Scale"}, scaleBox, new Label{ Text="Align"}, alignDrop, new Label{ Text="V-Align"}, vAlignDrop } },
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { gridCheck, wrapCheck, new Label{ Text="Text(mm)"}, textMm, new Label{ Text="Font"}, fontBox, new Label{ Text="DimStyle"}, dimStyleDrop } },
                new StackLayoutItem(new StackLayout
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Items = { pullButton, updateButton, applyButton }
                }, HorizontalAlignment.Stretch)
            }
        };

        updateButton.Click += (_, _) =>
        {
            // Bridge to command
            Rhino.RhinoApp.RunScript("_ExcelLinkUpdate", false);
        };

        applyButton.Click += (_, _) =>
        {
            var doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            var selected = doc.Objects.GetSelectedObjects(false, false).OfType<InstanceObject>().ToArray();
            foreach (var inst in selected)
            {
                var a = inst.Attributes.Duplicate();
                Set(a, Commands.BlockMetadataKeys.File, filePathText.Text);
                Set(a, Commands.BlockMetadataKeys.Sheet, sheetText.Text);
                Set(a, Commands.BlockMetadataKeys.Range, rangeText.Text);
                if (double.TryParse(scaleBox.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var sc))
                    Set(a, Commands.BlockMetadataKeys.Scale, sc.ToString(CultureInfo.InvariantCulture));
                var h = alignDrop.SelectedIndex switch { 1 => 0, 2 => 1, 3 => 2, _ => 0 };
                Set(a, Commands.BlockMetadataKeys.Align, h.ToString(CultureInfo.InvariantCulture));
                var v = vAlignDrop.SelectedIndex switch { 0 => 0, 2 => 2, _ => 1 };
                Set(a, Commands.BlockMetadataKeys.VAlign, v.ToString(CultureInfo.InvariantCulture));
                Set(a, Commands.BlockMetadataKeys.ShowGrid, gridCheck.Checked == true ? "1" : "0");
                if (double.TryParse(textMm.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var mm))
                    Set(a, Commands.BlockMetadataKeys.TextMm, mm.ToString(CultureInfo.InvariantCulture));
                else Set(a, Commands.BlockMetadataKeys.TextMm, string.Empty);
                Set(a, Commands.BlockMetadataKeys.Font, fontBox.Text ?? string.Empty);
                Set(a, Commands.BlockMetadataKeys.DimStyle, dimStyleDrop.SelectedValue?.ToString() ?? string.Empty);
                Set(a, Commands.BlockMetadataKeys.Wrap, wrapCheck.Checked == true ? "1" : "0");
                doc.Objects.ModifyAttributes(inst, a, true);
            }
            RhinoApp.RunScript("_ExcelLinkApplyAttributes", false);
        };

        // Load current selection into panel
        RhinoDoc.SelectObjects += (_, __) => LoadFromSelection(filePathText, sheetText, rangeText, scaleBox, alignDrop, vAlignDrop, gridCheck, wrapCheck, textMm, fontBox, dimStyleDrop);
        RhinoDoc.DeselectAllObjects += (_, __) => LoadFromSelection(filePathText, sheetText, rangeText, scaleBox, alignDrop, vAlignDrop, gridCheck, wrapCheck, textMm, fontBox, dimStyleDrop);
        RhinoDoc.DeselectObjects += (_, __) => LoadFromSelection(filePathText, sheetText, rangeText, scaleBox, alignDrop, vAlignDrop, gridCheck, wrapCheck, textMm, fontBox, dimStyleDrop);
        LoadFromSelection(filePathText, sheetText, rangeText, scaleBox, alignDrop, vAlignDrop, gridCheck, wrapCheck, textMm, fontBox, dimStyleDrop);
    }

    private static void Set(ObjectAttributes a, string key, string value) => a.SetUserString(key, value ?? string.Empty);

    private static void LoadFromSelection(TextBox file, TextBox sheet, TextBox range, TextBox scale, DropDown hAlign, DropDown vAlign, CheckBox grid, CheckBox wrap, TextBox tmm, TextBox font, DropDown dimStyle)
    {
        var doc = RhinoDoc.ActiveDoc; if (doc == null) return;
        var inst = doc.Objects.GetSelectedObjects(false, false).OfType<InstanceObject>().FirstOrDefault();
        if (inst == null) return;
        var def = inst.InstanceDefinition; if (def == null) return;
        string Read(string key)
            => !string.IsNullOrWhiteSpace(inst.Attributes.GetUserString(key)) ? inst.Attributes.GetUserString(key) : def.GetUserString(key);

        file.Text = Read(Commands.BlockMetadataKeys.File) ?? string.Empty;
        sheet.Text = Read(Commands.BlockMetadataKeys.Sheet) ?? string.Empty;
        range.Text = Read(Commands.BlockMetadataKeys.Range) ?? string.Empty;
        scale.Text = Read(Commands.BlockMetadataKeys.Scale) ?? string.Empty;
        if (int.TryParse(Read(Commands.BlockMetadataKeys.Align), out var h))
            hAlign.SelectedIndex = h switch { 0 => 1, 1 => 2, 2 => 3, _ => 0 };
        if (int.TryParse(Read(Commands.BlockMetadataKeys.VAlign), out var v))
            vAlign.SelectedIndex = v switch { 0 => 0, 2 => 2, _ => 1 };
        grid.Checked = Read(Commands.BlockMetadataKeys.ShowGrid) == "1";
        wrap.Checked = Read(Commands.BlockMetadataKeys.Wrap) == "1";
        tmm.Text = Read(Commands.BlockMetadataKeys.TextMm) ?? string.Empty;
        font.Text = Read(Commands.BlockMetadataKeys.Font) ?? string.Empty;
        var ds = Read(Commands.BlockMetadataKeys.DimStyle) ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(ds))
        {
            var list = dimStyle.DataStore?.Cast<string>().ToList();
            var idx = list?.FindIndex(n => string.Equals(n, ds, StringComparison.OrdinalIgnoreCase)) ?? -1;
            if (idx >= 0) dimStyle.SelectedIndex = idx;
        }
    }
}


