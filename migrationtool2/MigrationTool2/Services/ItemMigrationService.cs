using Microsoft.SharePoint.Client;
using MigrationTool2.Helpers;
using MigrationTool2.Models;

namespace MigrationTool2.Services;

public class ItemMigrationService
{
    private readonly ThrottleHandler _throttle;
    private readonly OptionsConfig _options;
    private readonly MigrationManifest _manifest;
    private readonly FieldHelper _fieldCopier;
    private readonly MigrationAuditLog _auditLog;

    public ItemMigrationService(
        ThrottleHandler throttle,
        OptionsConfig options,
        MigrationManifest manifest,
        FieldHelper fieldCopier,
        MigrationAuditLog auditLog)
    {
        _throttle = throttle;
        _options = options;
        _manifest = manifest;
        _fieldCopier = fieldCopier;
        _auditLog = auditLog;
    }

    public async Task MigrateAllItems(
        ClientContext sourceCtx, List sourceList,
        ClientContext destCtx, List destList,
        Dictionary<string, Field> destFields,
        string sourceListName, string destListName,
        MigrationResult result)
    {
        var pageSize = _options.BatchSize;
        int totalMigrated = 0;
        ListItemCollectionPosition? position = null;

        sourceCtx.Load(sourceList,
            l => l.Title, l => l.ItemCount,
            l => l.BaseType, l => l.BaseTemplate);
        sourceCtx.ExecuteQuery();

        var isDocLib = sourceList.BaseTemplate == (int)ListTemplateType.DocumentLibrary;
        if (isDocLib)
        {
            Console.WriteLine($"  '{sourceListName}' is a document library. Files handled by FileMigrationService.");
            return;
        }

        Console.WriteLine($"  Starting item migration for list '{sourceListName}' ({sourceList.ItemCount} items)...");

        var sourceFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var sourceFields = sourceList.Fields;
            sourceCtx.Load(sourceFields);
            sourceCtx.ExecuteQuery();
            foreach (var f in sourceFields)
                sourceFieldNames.Add(f.InternalName);
        }
        catch { }

        var destFieldTypes = destFields.ToDictionary(
            kv => kv.Key,
            kv => kv.Value.FieldTypeKind,
            StringComparer.OrdinalIgnoreCase);

        var filteredDestFields = destFields
            .Where(kv => sourceFieldNames.Contains(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

        var existingDestByTitle = new Dictionary<string, Queue<ListItem>>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var existingQuery = new CamlQuery
            {
                ViewXml = @"<View><ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/></ViewFields><RowLimit>5000</RowLimit></View>"
            };
            var existingItems = destList.GetItems(existingQuery);
            destCtx.Load(existingItems);
            destCtx.ExecuteQuery();
            foreach (var item in existingItems)
            {
                try
                {
                    var t = item["Title"]?.ToString();
                    if (!string.IsNullOrEmpty(t))
                    {
                        if (!existingDestByTitle.ContainsKey(t))
                            existingDestByTitle[t] = new Queue<ListItem>();
                        existingDestByTitle[t].Enqueue(item);
                    }
                }
                catch { }
            }
        }
        catch { }

        Console.WriteLine($"  Found {existingDestByTitle.Values.Sum(q => q.Count)} existing items on destination.");

        while (true)
        {
            var sysFields = _options.PreserveCreatedBy || _options.PreserveModifiedBy || _options.PreserveCreated || _options.PreserveModified
                ? "<FieldRef Name='Created'/><FieldRef Name='Modified'/><FieldRef Name='Author'/><FieldRef Name='Editor'/>"
                : "";
            var fieldRefs = string.Join("",
                filteredDestFields.Keys.Select(f => $"<FieldRef Name='{f}'/>"));

            var query = new CamlQuery
            {
                ListItemCollectionPosition = position,
                ViewXml = $@"<View Scope='RecursiveAll'>
                    <ViewFields>
                        <FieldRef Name='ID'/>
                        <FieldRef Name='Title'/>
                        <FieldRef Name='ContentTypeId'/>
                        {fieldRefs}
                        {sysFields}
                    </ViewFields>
                    <RowLimit>{pageSize}</RowLimit>
                </View>"
            };

            var items = sourceList.GetItems(query);
            sourceCtx.Load(items);
            sourceCtx.ExecuteQuery();

            if (items.Count == 0) break;

            position = items.ListItemCollectionPosition;

            try
            {
                foreach (var item in items)
                    sourceCtx.Load(item, i => i.FieldValuesAsText);
                sourceCtx.ExecuteQuery();
            }
            catch { }

            foreach (var sourceItem in items)
            {
                var title = GetItemTitle(sourceItem);

                if (_options.OverwriteMode == "Skip" && _manifest.IsAlreadyMigrated(sourceListName, sourceItem.Id))
                {
                    result.Skipped++;
                    _auditLog.Record(AuditAction.ItemSkipped, sourceListName, destListName,
                        srcId: sourceItem.Id.ToString(),
                        srcName: title,
                        status: "Skipped", detail: "OverwriteMode=Skip, already migrated");
                    continue;
                }

                try
                {
                    existingDestByTitle.TryGetValue(title, out var existingQueue);
                    ListItem? existingItem = existingQueue?.Count > 0 ? existingQueue.Dequeue() : null;

                    await _throttle.ExecuteWithRetryAsync(async () =>
                    {
                        var (destItem, destId) = _fieldCopier.CreateAndCopyFields(
                            destList, sourceItem, sourceCtx, destCtx, filteredDestFields,
                            sourceListName, destListName, existingItem);

                        var action = existingItem != null ? AuditAction.ItemUpdated : AuditAction.ItemCreated;
                        _manifest.RecordItem(sourceListName, sourceItem.Id, destId, title, "OK");
                        result.Succeeded++;
                        totalMigrated++;

                        _auditLog.Record(action, sourceListName, destListName,
                            srcId: sourceItem.Id.ToString(), destId: destId.ToString(),
                            srcName: title, destName: title);

                        Console.WriteLine($"  {(existingItem != null ? "~" : "+")} [{totalMigrated}] {title}{(existingItem != null ? " (overwritten)" : "")}");
                    }, $"Item: {title} (ID: {sourceItem.Id})");
                }
                catch (Exception ex)
                {
                    result.Failed++;
                    result.Errors.Add(new MigrationError
                    {
                        ListName = sourceListName,
                        ItemId = sourceItem.Id.ToString(),
                        Title = title,
                        Message = ex.Message,
                        Exception = ex
                    });

                    _auditLog.Record(AuditAction.ItemFailed, sourceListName, destListName,
                        srcId: sourceItem.Id.ToString(),
                        srcName: title,
                        status: "Failed", detail: ex.Message);

                    Console.WriteLine($"  ! FAILED: {title} - {ex.Message}");
                }
            }

            if (position == null) break;
            Console.WriteLine($"  Page complete. Progress: {totalMigrated}/{sourceList.ItemCount}");
        }

        Console.WriteLine($"  Item migration complete. Total migrated: {totalMigrated}");
    }

    private static string GetItemTitle(ListItem item)
    {
        try
        {
            var t = item["Title"];
            if (t != null) return t.ToString()!;
        }
        catch { }

        try
        {
            var t = item.FieldValuesAsText["Title"];
            if (!string.IsNullOrEmpty(t)) return t;
        }
        catch { }

        return item.Id.ToString();
    }
}
