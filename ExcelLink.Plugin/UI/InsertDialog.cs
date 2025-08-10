using System;
using System.IO;
using Eto.Forms;
using Eto.Drawing;

namespace ExcelLink.Plugin.UI;

public sealed class InsertDialog : Dialog<InsertDialog.ResultModel>
{
    public sealed class ResultModel
    {
        public string FilePath { get; init; } = string.Empty;
        public string Sheet { get; init; } = string.Empty;
        public string RangeOrName { get; init; } = string.Empty;
        public double Scale { get; init; } = 1.0;
        public string Alignment { get; init; } = "Excel"; // Excel|Left|Center|Right
        public bool DrawGrid { get; init; } = true;
        public double? TextHeightMm { get; init; }
        public string? FontFamily { get; init; }
        public bool Wrap { get; init; }
        public string? DimStyleName { get; init; }
    }

    private readonly TextBox _file = new() { Width = 380 };
    private readonly ComboBox _sheet = new() { Width = 200 };
    private readonly TextBox _range = new() { Width = 160 };
    private readonly NumericStepper _scale = new() { Value = 1.0, DecimalPlaces = 3, MinValue = 0.001, MaxValue = 1000 };
    private readonly DropDown _align = new() { DataStore = new[] { "Excel", "Left", "Center", "Right" } };
    private readonly CheckBox _grid = new() { Text = "Grid", Checked = true };
    private readonly NumericStepper _textHeight = new() { DecimalPlaces = 2, MinValue = 0, MaxValue = 100 };
    private readonly CheckBox _heightFromDimStyle = new() { Text = "Height from DimStyle", Checked = true };
    private readonly TextBox _font = new() { PlaceholderText = "Font family (optional)", Width = 200 };
    private readonly CheckBox _wrap = new() { Text = "Wrap", Checked = false };
    private readonly ImageView _preview = new() { Size = new Size(280, 160) };
    private readonly DropDown _dimStyle = new() { Width = 220 };

    public InsertDialog()
    {
        Title = "Insert Excel Link";
        Resizable = false;
        Padding = 10;

        var pick = new Button { Text = "Browse…" };
        pick.Click += (_, _) =>
        {
        var ofd = new OpenFileDialog { MultiSelect = false, Filters = { new FileFilter("Excel", ".xlsx") } };
            if (ofd.ShowDialog(this) == DialogResult.Ok)
        {
            _file.Text = ofd.FileName;
            TryPopulateSheetsAndRange(ofd.FileName);
        }
        };

        var main = new StackLayout
        {
            Spacing = 8,
            Items =
            {
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { new Label{ Text = "File"}, _file, pick } },
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { new Label{ Text = "Sheet"}, _sheet, new Label{ Text = "Range/Name"}, _range } },
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { new Label{ Text = "Scale"}, _scale, new Label{ Text = "Align"}, _align } },
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { _grid, new Label{ Text = "Text (mm)"}, _textHeight, _heightFromDimStyle, _wrap, _font } },
                new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { new Label{ Text = "DimStyle"}, _dimStyle } },
                new GroupBox { Text = "Preview", Content = _preview },
            }
        };

        DefaultButton = new Button { Text = "OK" };
        AbortButton = new Button { Text = "Cancel" };
        DefaultButton.Click += (_, _) => OnOk();
        AbortButton.Click += (_, _) => Close();

        var buttons = new StackLayout { Orientation = Orientation.Horizontal, Spacing = 6, Items = { DefaultButton, AbortButton } };
        Content = new StackLayout { Spacing = 10, Items = { main, buttons } };

        PopulateDimStyles();
        _sheet.SelectedIndexChanged += (_, _) => { if (!string.IsNullOrWhiteSpace(_file.Text) && _sheet.SelectedValue is string sname) { var used = ExcelLink.Core.Services.ExcelInspector.GetUsedRangeA1(_file.Text, sname); if (!string.IsNullOrWhiteSpace(used)) _range.Text = used; } UpdatePreview(); };
        _range.TextChanged += (_, _) => UpdatePreview();
        _grid.CheckedChanged += (_, _) => UpdatePreview();
        _textHeight.ValueChanged += (_, _) => UpdatePreview();
        _font.TextChanged += (_, _) => UpdatePreview();
        _wrap.CheckedChanged += (_, _) => UpdatePreview();
        UpdatePreview();
    }

    private void TryPopulateSheetsAndRange(string filePath)
    {
        try
        {
            var sheets = ExcelLink.Core.Services.ExcelInspector.ListSheetNames(filePath);
            _sheet.DataStore = sheets;
            if (sheets.Count > 0)
            {
                _sheet.SelectedIndex = 0;
                var used = ExcelLink.Core.Services.ExcelInspector.GetUsedRangeA1(filePath, sheets[0]);
                if (!string.IsNullOrWhiteSpace(used))
                    _range.Text = used;
            }
        }
        catch
        {
            // ignore UI population errors
        }
    }

    // Basic bitmap preview (approximate)
    private void UpdatePreview()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_file.Text) || _sheet.SelectedValue == null || string.IsNullOrWhiteSpace(_range.Text))
            { _preview.Image = null; return; }
            var file = _file.Text;
            var sheet = _sheet.SelectedValue.ToString();
            var range = _range.Text;
            var reader = new ExcelLink.Core.Rendering.ExcelClosedXmlReader();
            var table = reader.ReadTable(file, sheet, range);
            var bmp = new Bitmap(_preview.Size.Width, _preview.Size.Height, PixelFormat.Format32bppRgba);
            using (var g = new Graphics(bmp))
            {
                g.Clear(Colors.White);
                int rows = Math.Min(table.Rows.Count, 20);
                int cols = Math.Min(table.Columns.Count, 10);
                float w = _preview.Size.Width - 10;
                float h = _preview.Size.Height - 10;
                float x0 = 5, y0 = 5;
                float cw = w / Math.Max(1, cols);
                float rh = h / Math.Max(1, rows);
                var pen = new Pen(Color.FromArgb(200, 200, 200), 1);
                if (_grid.Checked == true)
                {
                    for (int r = 0; r <= rows; r++) g.DrawLine(pen, x0, y0 + r * rh, x0 + cols * cw, y0 + r * rh);
                    for (int c = 0; c <= cols; c++) g.DrawLine(pen, x0 + c * cw, y0, x0 + c * cw, y0 + rows * rh);
                }
                var fontName = string.IsNullOrWhiteSpace(_font.Text) ? SystemFonts.Default().FamilyName : _font.Text;
                var font = new Font(fontName, 8);
                for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols && c < table.Rows[r].Cells.Count; c++)
                {
                    var text = table.Rows[r].Cells[c].Value?.ToString() ?? string.Empty;
                    var pt = new PointF(x0 + c * cw + 2, y0 + r * rh + 2);
                    g.DrawText(font, Colors.Black, pt, text);
                }
            }
            _preview.Image = bmp;
        }
        catch { _preview.Image = null; }
    }

    private void OnOk()
    {
        if (string.IsNullOrWhiteSpace(_file.Text) || !File.Exists(_file.Text))
        {
            MessageBox.Show(this, "Please choose an existing .xlsx file.", MessageBoxType.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(_sheet.Text) || string.IsNullOrWhiteSpace(_range.Text))
        {
            MessageBox.Show(this, "Please enter sheet and range or named range.", MessageBoxType.Warning);
            return;
        }

        Result = new ResultModel
        {
            FilePath = _file.Text,
            Sheet = _sheet.Text,
            RangeOrName = _range.Text,
            Scale = _scale.Value,
            Alignment = _align.SelectedValue?.ToString() ?? "Excel",
            DrawGrid = _grid.Checked == true,
            TextHeightMm = _heightFromDimStyle.Checked == true ? null : (_textHeight.Value > 0 ? _textHeight.Value : null),
            FontFamily = string.IsNullOrWhiteSpace(_font.Text) ? null : _font.Text,
            Wrap = _wrap.Checked == true,
            DimStyleName = _dimStyle.SelectedValue?.ToString()
        };
        Close(Result);
    }

    private void PopulateDimStyles()
    {
        try
        {
            var doc = Rhino.RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                _dimStyle.DataStore = new[] { "(none)" }; _dimStyle.SelectedIndex = 0; return;
            }
            var names = doc.DimStyles.Select(ds => ds.Name).ToList();
            if (names.Count == 0) names.Add("(none)");
            _dimStyle.DataStore = names; _dimStyle.SelectedIndex = 0;
        }
        catch
        {
            _dimStyle.DataStore = new[] { "(none)" }; _dimStyle.SelectedIndex = 0;
        }
    }
}


