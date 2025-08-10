using System;
using System.Runtime.InteropServices;
using Eto.Drawing;
using Eto.Forms;

namespace ExcelLink.Plugin.UI;

[Guid("6ddcb0c5-5fb1-4b9a-acdf-9b7e7a6d3a03")]
public sealed class ExcelLinkPanel : Eto.Forms.Panel
{
    public ExcelLinkPanel()
    {
        var filePathText = new TextBox { PlaceholderText = "Excel file path" };
        var sheetText = new TextBox { PlaceholderText = "Sheet" };
        var rangeText = new TextBox { PlaceholderText = "Range or Named Range" };

        var pullButton = new Button { Text = "Pull (Excel → Rhino)" };
        var updateButton = new Button { Text = "Update Selected" };

        Content = new StackLayout
        {
            Padding = 10,
            Spacing = 6,
            Items =
            {
                new Label { Text = "ExcelLink Panel", Font = new Font(SystemFont.Bold, 12) },
                filePathText,
                sheetText,
                rangeText,
                new StackLayoutItem(new StackLayout
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Items = { pullButton, updateButton }
                }, HorizontalAlignment.Stretch)
            }
        };

        updateButton.Click += (_, _) =>
        {
            // Bridge to command
            Rhino.RhinoApp.RunScript("_ExcelLinkUpdate", false);
        };
    }
}


