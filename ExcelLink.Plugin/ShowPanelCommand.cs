using System;
using System.Runtime.InteropServices;
using Rhino;
using Rhino.Commands;
using Rhino.UI;

namespace ExcelLink.Plugin;

[Guid("c1b7a7df-8e15-4c8c-8ba4-2e5f0c51efb2")]
public sealed class ExcelLinkShowPanelCommand : Command
{
    public override string EnglishName => "ExcelLinkPanel";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var panelId = typeof(UI.ExcelLinkPanel).GUID;
        Panels.OpenPanel(panelId);
        return Result.Success;
    }
}


