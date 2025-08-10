using System.Reflection;
using System.Runtime.InteropServices;
using Rhino.PlugIns;

// Required plug-in metadata for Rhino
[assembly: PlugInDescription(DescriptionType.Address, "")]
[assembly: PlugInDescription(DescriptionType.Country, "")]
[assembly: PlugInDescription(DescriptionType.Email, "")]
[assembly: PlugInDescription(DescriptionType.Phone, "")]
[assembly: PlugInDescription(DescriptionType.Fax, "")]
[assembly: PlugInDescription(DescriptionType.Organization, "ExcelLink Team")]
[assembly: PlugInDescription(DescriptionType.UpdateUrl, "https://example.com")] 
[assembly: AssemblyTitle("ExcelLink")] 
[assembly: AssemblyDescription("Insert, update, and roundtrip Excel tables in Rhino.")]

// Assembly-level plug-in Id (required by Rhino)
[assembly: Guid("8f1a9b17-3f57-4b6a-9b8e-55c43a64d0a3")]

