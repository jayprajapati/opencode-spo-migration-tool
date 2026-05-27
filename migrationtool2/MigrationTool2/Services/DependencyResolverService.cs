using Microsoft.SharePoint.Client;

namespace MigrationTool2.Services;

public class DependencyResolverService
{
    public class ListDependency
    {
        public string SourceListName { get; set; } = string.Empty;
        public string TargetListName { get; set; } = string.Empty;
        public string? TargetListOnDestination { get; set; }
        public List<string> LookupFieldNames { get; set; } = new();
        public string DependsOnList { get; set; } = string.Empty;
        public string DependsOnEntity { get; set; } = string.Empty;
    }

    public class LookupFieldInfo
    {
        public string InternalName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string DependsOnList { get; set; } = string.Empty;
        public string DependsOnField { get; set; } = string.Empty;
        public Guid LookupListId { get; set; }
        public bool IsMulti { get; set; }
    }

    public List<LookupFieldInfo> GetLookupFields(ClientContext sourceCtx, string sourceListName)
    {
        var sourceList = sourceCtx.Web.Lists.GetByTitle(sourceListName);
        sourceCtx.Load(sourceList, l => l.Fields);
        sourceCtx.ExecuteQuery();

        var lookups = new List<LookupFieldInfo>();
        foreach (var field in sourceList.Fields)
        {
            if (field.FieldTypeKind != FieldType.Lookup) continue;

            var lookupField = sourceCtx.CastTo<FieldLookup>(field);
            sourceCtx.Load(lookupField,
                f => f.LookupList, f => f.LookupField,
                f => f.Title, f => f.InternalName,
                f => f.AllowMultipleValues);
            sourceCtx.ExecuteQuery();

            if (string.IsNullOrEmpty(lookupField.LookupList)) continue;
            if (!Guid.TryParse(lookupField.LookupList, out var lookupListGuid)) continue;

            try
            {
                var lookupList = sourceCtx.Web.Lists.GetById(lookupListGuid);
                sourceCtx.Load(lookupList, l => l.Title);
                sourceCtx.ExecuteQuery();

                lookups.Add(new LookupFieldInfo
                {
                    InternalName = lookupField.InternalName,
                    Title = lookupField.Title,
                    DependsOnList = lookupList.Title,
                    DependsOnField = lookupField.LookupField,
                    LookupListId = lookupListGuid,
                    IsMulti = lookupField.AllowMultipleValues
                });
            }
            catch
            {
                Console.WriteLine($"  [WARN] Could not resolve lookup target for field '{lookupField.Title}'. Skipping dependency.");
            }
        }

        return lookups;
    }

    public List<string> ResolveMigrationOrder(
        ClientContext sourceCtx,
        List<string> requestedLists,
        bool autoResolveDependencies)
    {
        var allLists = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var edges = new List<(string from, string to)>();

        foreach (var listName in requestedLists)
        {
            allLists.Add(listName);
            if (!autoResolveDependencies) continue;

            var lookups = GetLookupFields(sourceCtx, listName);
            foreach (var lookup in lookups)
            {
                var depList = lookup.DependsOnList;
                if (string.IsNullOrEmpty(depList)) continue;

                edges.Add((depList, listName));
                allLists.Add(depList);
            }
        }

        var sorted = TopologicalSort(allLists.ToList(), edges);

        var autoIncluded = sorted
            .Where(s => !requestedLists.Contains(s, StringComparer.OrdinalIgnoreCase))
            .ToList();

        if (autoIncluded.Count > 0)
        {
            Console.WriteLine($"  Auto-included dependency lists: {string.Join(", ", autoIncluded)}");
        }

        return sorted;
    }

    public List<string> TopologicalSort(List<string> allLists, List<(string from, string to)> edges)
    {
        var graph = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var inDegree = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var list in allLists)
        {
            graph[list] = new List<string>();
            inDegree[list] = 0;
        }

        foreach (var (from, to) in edges)
        {
            if (!graph.ContainsKey(from) || !graph.ContainsKey(to)) continue;
            if (string.Equals(from, to, StringComparison.OrdinalIgnoreCase)) continue;

            graph[from].Add(to);
            inDegree[to] = inDegree.GetValueOrDefault(to) + 1;
        }

        var queue = new Queue<string>();
        foreach (var (node, degree) in inDegree)
        {
            if (degree == 0) queue.Enqueue(node);
        }

        var sorted = new List<string>();
        while (queue.Count > 0)
        {
            var node = queue.Dequeue();
            sorted.Add(node);

            if (!graph.ContainsKey(node)) continue;
            foreach (var neighbor in graph[node])
            {
                inDegree[neighbor]--;
                if (inDegree[neighbor] == 0) queue.Enqueue(neighbor);
            }
        }

        foreach (var list in allLists)
        {
            if (!sorted.Contains(list, StringComparer.OrdinalIgnoreCase))
                sorted.Add(list);
        }

        return sorted;
    }

    public Dictionary<string, string> BuildListNameMapping(
        ClientContext sourceCtx,
        List<string> sourceListNames,
        ClientContext destCtx)
    {
        var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var srcName in sourceListNames)
        {
            try
            {
                var srcList = sourceCtx.Web.Lists.GetByTitle(srcName);
                sourceCtx.Load(srcList, l => l.Title, l => l.EntityTypeName);
                sourceCtx.ExecuteQuery();

                try
                {
                    var destList = destCtx.Web.Lists.GetByTitle(srcName);
                    destCtx.Load(destList, l => l.Title, l => l.Id);
                    destCtx.ExecuteQuery();
                    mapping[srcList.Title] = destList.Title;
                }
                catch
                {
                    mapping[srcList.Title] = srcName;
                }
            }
            catch
            {
                mapping[srcName] = srcName;
            }
        }

        return mapping;
    }
}
