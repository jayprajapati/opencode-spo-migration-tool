using System.Globalization;

namespace MigrationTool2.Models;

public class MigrationManifest
{
    private readonly string _filePath;
    private readonly Dictionary<string, Dictionary<int, int>> _listIdMappings = new();
    private readonly Dictionary<string, HashSet<int>> _migratedSourceIds = new();
    private readonly object _lock = new();

    public MigrationManifest(string filePath)
    {
        _filePath = filePath;
    }

    public bool IsAlreadyMigrated(string listName, int sourceItemId)
    {
        lock (_lock)
        {
            return _migratedSourceIds.TryGetValue(listName, out var ids) && ids.Contains(sourceItemId);
        }
    }

    public void RecordItem(string listName, int sourceItemId, int destItemId, string title, string status)
    {
        lock (_lock)
        {
            if (!_listIdMappings.ContainsKey(listName))
                _listIdMappings[listName] = new Dictionary<int, int>();
            _listIdMappings[listName][sourceItemId] = destItemId;

            if (!_migratedSourceIds.ContainsKey(listName))
                _migratedSourceIds[listName] = new HashSet<int>();
            _migratedSourceIds[listName].Add(sourceItemId);
        }

        var line = $"{EscapeCsv(listName)},{sourceItemId},{destItemId},{EscapeCsv(title)},{status},{DateTime.UtcNow:O}";
        File.AppendAllLines(_filePath, new[] { line });
    }

    public int? GetDestinationId(string listName, int sourceId)
    {
        lock (_lock)
        {
            if (_listIdMappings.TryGetValue(listName, out var mapping) && mapping.TryGetValue(sourceId, out var destId))
                return destId;
        }
        return null;
    }

    public void LoadExisting()
    {
        if (!File.Exists(_filePath)) return;

        foreach (var line in File.ReadLines(_filePath))
        {
            var parts = SplitCsvLine(line);
            if (parts.Length < 5) continue;

            var listName = parts[0];
            if (!int.TryParse(parts[1], out var srcId)) continue;
            if (!int.TryParse(parts[2], out var destId)) continue;

            if (!_listIdMappings.ContainsKey(listName))
                _listIdMappings[listName] = new Dictionary<int, int>();
            _listIdMappings[listName][srcId] = destId;

            if (!_migratedSourceIds.ContainsKey(listName))
                _migratedSourceIds[listName] = new HashSet<int>();
            _migratedSourceIds[listName].Add(srcId);
        }
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }

    private static string[] SplitCsvLine(string line)
    {
        var result = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }
        result.Add(current.ToString());
        return result.ToArray();
    }

    public IReadOnlyDictionary<string, IReadOnlyDictionary<int, int>> GetAllMappings()
    {
        lock (_lock)
        {
            return _listIdMappings
                .ToDictionary(kv => kv.Key, kv => (IReadOnlyDictionary<int, int>)kv.Value);
        }
    }
}
