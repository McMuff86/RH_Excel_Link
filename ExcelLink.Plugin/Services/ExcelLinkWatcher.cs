using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Rhino;
using Rhino.DocObjects;
using Rhino.UI;

namespace ExcelLink.Plugin.Services;

internal sealed class ExcelLinkWatcherService
{
    private sealed class WatchEntry
    {
        public FileSystemWatcher Watcher = null!;
        public HashSet<int> DefinitionIndices = new();
        public DateTime LastEventUtc = DateTime.MinValue;
        public string FileName = string.Empty;
    }

    private static readonly Lazy<ExcelLinkWatcherService> LazyInstance = new(() => new ExcelLinkWatcherService());
    public static ExcelLinkWatcherService Instance => LazyInstance.Value;

    private readonly Dictionary<string, WatchEntry> _filePathToEntry = new(StringComparer.OrdinalIgnoreCase);

    private ExcelLinkWatcherService() { }

    public void RegisterDefinition(RhinoDoc doc, int defIndex, string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !Path.IsPathRooted(filePath)) return;
        var dir = Path.GetDirectoryName(filePath);
        var name = Path.GetFileName(filePath);
        if (string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(name)) return;

        if (!_filePathToEntry.TryGetValue(filePath, out var entry))
        {
            var fsw = new FileSystemWatcher(dir, name)
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size
            };
            entry = new WatchEntry
            {
                Watcher = fsw,
                FileName = name
            };
            fsw.Changed += (_, __) => OnFileChanged(doc, filePath);
            fsw.Renamed += (_, __) => OnFileChanged(doc, filePath);
            fsw.Deleted += (_, __) => OnFileChanged(doc, filePath);
            fsw.EnableRaisingEvents = true;
            _filePathToEntry[filePath] = entry;
        }

        entry.DefinitionIndices.Add(defIndex);
    }

    public void UnregisterDefinition(int defIndex, string filePath)
    {
        if (!_filePathToEntry.TryGetValue(filePath, out var entry)) return;
        entry.DefinitionIndices.Remove(defIndex);
        if (entry.DefinitionIndices.Count == 0)
        {
            entry.Watcher.EnableRaisingEvents = false;
            entry.Watcher.Dispose();
            _filePathToEntry.Remove(filePath);
        }
    }

    public void UpdateDefinitionFile(RhinoDoc doc, int defIndex, string oldFilePath, string newFilePath)
    {
        if (!string.IsNullOrWhiteSpace(oldFilePath))
            UnregisterDefinition(defIndex, oldFilePath);
        if (!string.IsNullOrWhiteSpace(newFilePath))
            RegisterDefinition(doc, defIndex, newFilePath);
    }

    public void ReplaceDefinitionIndex(string filePath, int oldIndex, int newIndex)
    {
        if (!_filePathToEntry.TryGetValue(filePath, out var entry)) return;
        if (entry.DefinitionIndices.Remove(oldIndex))
            entry.DefinitionIndices.Add(newIndex);
    }

    private void OnFileChanged(RhinoDoc doc, string filePath)
    {
        if (!_filePathToEntry.TryGetValue(filePath, out var entry)) return;
        // debounce
        var now = DateTime.UtcNow;
        if ((now - entry.LastEventUtc).TotalMilliseconds < 500) return;
        entry.LastEventUtc = now;

        RhinoApp.InvokeOnUiThread(() =>
        {
            var res = Dialogs.ShowMessage(
                $"Excel file changed:\n{filePath}\n\nUpdate linked tables now? (Yes)\nExtend ranges to used range? (Yes = extend)",
                "ExcelLink",
                ShowMessageButton.YesNoCancel,
                ShowMessageIcon.Question);
            if (res == ShowMessageResult.Cancel)
                return;

            var indices = entry.DefinitionIndices.ToArray();
            foreach (var defIndex in indices)
            {
                if (defIndex < 0 || defIndex >= doc.InstanceDefinitions.Count)
                    continue;
                var def = doc.InstanceDefinitions[defIndex];
                if (def == null) continue;
                try
                {
                    bool extend = res == ShowMessageResult.Yes;
                    ExcelLink.Plugin.Commands.ExcelLinkUpdater.UpdateDefinition(doc, def, extendRangeIfNeeded: extend);
                }
                catch (Exception ex)
                {
                    RhinoApp.WriteLine($"Auto-update failed for '{def.Name}': {ex.Message}");
                }
            }

            doc.Views.Redraw();
        });
    }
}


