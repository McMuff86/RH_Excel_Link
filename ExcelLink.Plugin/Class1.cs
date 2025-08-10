using System;
using System.Runtime.InteropServices;
using Rhino.PlugIns;
using Rhino.UI;

namespace ExcelLink.Plugin;

[Guid("8f1a9b17-3f57-4b6a-9b8e-55c43a64d0a3")]
public sealed class ExcelLinkPlugin : PlugIn
{
    public static ExcelLinkPlugin Instance { get; private set; } = null!;

    public ExcelLinkPlugin()
    {
        Instance = this;
    }

    protected override LoadReturnCode OnLoad(ref string errorMessage)
    {
        Panels.RegisterPanel(this, typeof(UI.ExcelLinkPanel), "ExcelLink", null);
        return LoadReturnCode.Success;
    }
}
