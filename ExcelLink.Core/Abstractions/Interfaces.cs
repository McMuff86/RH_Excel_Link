using ExcelLink.Core.Models;

namespace ExcelLink.Core.Abstractions;

public interface IExcelReader
{
    TableModel ReadTable(string filePath, string sheetName, string rangeOrNamedRange);
}

public interface IExcelWriter
{
    void WriteTable(string filePath, string sheetName, string rangeOrNamedRange, TableModel table);
}

public interface ITableNormalizer
{
    TableModel Normalize(TableModel input);
}

public interface ITableRenderer
{
    // Rhino-dependent implementation will live in the plugin layer; this is a placeholder abstraction.
}

public interface ILinkMetadataStore
{
    // Placeholder for block metadata persistence abstraction; implemented in plugin project.
}


