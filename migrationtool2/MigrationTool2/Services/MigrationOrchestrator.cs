using System.Net;
using Microsoft.SharePoint.Client;
using MigrationTool2.Helpers;
using MigrationTool2.Models;

namespace MigrationTool2.Services;

public class MigrationOrchestrator
{
    private readonly MigrationConfig _config;
    private readonly AuthService _authService;
    private readonly ListProvisioningService _listProvisioner;
    private readonly FileMigrationService _fileMigrator;
    private readonly ItemMigrationService _itemMigrator;
    private readonly AttachmentMigrationService _attachmentMigrator;
    private readonly ContentTypeMigrationService _contentTypeMigrator;
    private readonly DependencyResolverService _dependencyResolver;
    private readonly MigrationManifest _manifest;
    private readonly MigrationAuditLog _auditLog;
    private readonly FieldHelper _fieldCopier;

    public MigrationOrchestrator(
        MigrationConfig config,
        AuthService authService,
        ListProvisioningService listProvisioner,
        FileMigrationService fileMigrator,
        ItemMigrationService itemMigrator,
        AttachmentMigrationService attachmentMigrator,
        ContentTypeMigrationService contentTypeMigrator,
        DependencyResolverService dependencyResolver,
        MigrationManifest manifest,
        MigrationAuditLog auditLog,
        FieldHelper fieldCopier)
    {
        _config = config;
        _authService = authService;
        _listProvisioner = listProvisioner;
        _fileMigrator = fileMigrator;
        _itemMigrator = itemMigrator;
        _attachmentMigrator = attachmentMigrator;
        _contentTypeMigrator = contentTypeMigrator;
        _dependencyResolver = dependencyResolver;
        _manifest = manifest;
        _auditLog = auditLog;
        _fieldCopier = fieldCopier;
    }

    public async Task<MigrationResult> ExecuteAsync()
    {
        var result = new MigrationResult
        {
            StartTime = DateTime.Now
        };

        var listsToMigrate = ResolveListOrder();

        Console.WriteLine("\n" + new string('=', 50));
        Console.WriteLine("  SharePoint Migration Tool v2");
        Console.WriteLine(new string('=', 50));
        Console.WriteLine($"  Source:      {_config.SourceSiteUrl}");
        Console.WriteLine($"  Destination: {_config.DestinationSiteUrl}");
        Console.WriteLine($"  Auth Mode:   {_config.Auth.Mode}");
        Console.WriteLine($"  Overwrite:   {_config.Options.OverwriteMode}");
        Console.WriteLine($"  Lists:       {string.Join(", ", listsToMigrate)}");
        Console.WriteLine(new string('=', 50) + "\n");

        Console.WriteLine("[1/5] Connecting to SharePoint...");
        var sourceCtx = _authService.GetContext(_config.SourceSiteUrl);
        var destCtx = _authService.GetContext(_config.DestinationSiteUrl);

        try
        {
            sourceCtx.ExecuteQuery();
            Console.WriteLine("  Source OK.");
        }
        catch (WebException ex) when (ex.Response is HttpWebResponse resp)
        {
            using var reader = new StreamReader(resp.GetResponseStream());
            var body = reader.ReadToEnd();
            Console.WriteLine($"  Source 401: {body}");
            throw;
        }

        try
        {
            destCtx.ExecuteQuery();
            Console.WriteLine("  Dest OK.");
        }
        catch (WebException ex) when (ex.Response is HttpWebResponse resp)
        {
            using var reader = new StreamReader(resp.GetResponseStream());
            var body = reader.ReadToEnd();
            Console.WriteLine($"  Dest 401: {body}");
            throw;
        }

        Console.WriteLine("  Connected successfully.\n");

        var allDeferredLookups = new Dictionary<string, List<ListProvisioningService.DeferredLookupField>>(StringComparer.OrdinalIgnoreCase);
        var destListNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var listName in listsToMigrate)
        {
            Console.WriteLine($"\n{'='}{'='}{'='}{'='}{'='} Processing: '{listName}' {'='}{'='}{'='}{'='}{'='}\n");

            try
            {
                await MigrateSingleList(sourceCtx, destCtx, listName, result, allDeferredLookups, destListNames);
            }
            catch (Exception ex)
            {
                var detail = $"[{ex.GetType().Name}] {ex.Message}";
                Console.WriteLine($"\n  ! FAILED '{listName}': {detail}");
                result.ListErrors.Add($"{listName}: {detail}");
                _auditLog.Record(AuditAction.ListProvisionFailed, listName, listName,
                    status: "Failed", detail: detail);
            }
        }

        Console.WriteLine("\n[5/5] Finalizing...");
        result.EndTime = DateTime.Now;
        result.TotalItems = result.Succeeded + result.Failed + result.Skipped;

        Console.WriteLine(new string('=', 50));
        Console.WriteLine("  Migration Complete");
        Console.WriteLine(new string('=', 50));
        result.PrintSummary();

        return result;
    }

    private async Task MigrateSingleList(
        ClientContext sourceCtx, ClientContext destCtx,
        string listName, MigrationResult result,
        Dictionary<string, List<ListProvisioningService.DeferredLookupField>> allDeferredLookups,
        Dictionary<string, string> destListNames)
    {
        Console.WriteLine("[2/5] Loading source list...");
        var sourceList = sourceCtx.Web.Lists.GetByTitle(listName);
        sourceCtx.Load(sourceList,
            l => l.Title, l => l.ItemCount, l => l.BaseTemplate,
            l => l.BaseType, l => l.RootFolder.ServerRelativeUrl, l => l.Fields,
            l => l.Hidden, l => l.ContentTypesEnabled);
        sourceCtx.ExecuteQuery();

        var destListName = GetDestListName(listName);
        Console.WriteLine($"  Source: '{sourceList.Title}' ({sourceList.ItemCount} items)");
        Console.WriteLine($"  Dest:   '{destListName}'\n");

        Console.WriteLine("[3/5] Provisioning destination list...");
        var (destList, deferredLookups) = _listProvisioner.GetOrCreateList(
            destCtx, destListName, sourceList, sourceCtx);

        destListNames[listName] = destListName;
        allDeferredLookups[listName] = deferredLookups;

        var destFields = _listProvisioner.GetFieldMap(destList, destCtx);

        destCtx.Load(destList, l => l.RootFolder.ServerRelativeUrl);
        destCtx.ExecuteQuery();

        if (_config.Options.MigrateContentTypes)
        {
            Console.WriteLine("  Migrating content types...");
            _contentTypeMigrator.MigrateContentTypes(sourceCtx, sourceList, destCtx, destList);
        }

        result.CurrentList = listName;

        var isDocLib = sourceList.BaseTemplate == (int)ListTemplateType.DocumentLibrary;

        if (isDocLib)
        {
            Console.WriteLine("[4/5] Migrating files...\n");
            var sourceRootFolder = sourceCtx.Web.GetFolderByServerRelativeUrl(sourceList.RootFolder.ServerRelativeUrl);
            var destRootFolder = destCtx.Web.GetFolderByServerRelativeUrl(destList.RootFolder.ServerRelativeUrl);
            sourceCtx.Load(sourceRootFolder, f => f.ServerRelativeUrl, f => f.Folders, f => f.Files);
            destCtx.Load(destRootFolder, f => f.ServerRelativeUrl, f => f.Folders, f => f.Files);
            sourceCtx.ExecuteQuery();
            destCtx.ExecuteQuery();

            await _fileMigrator.MigrateFilesInFolder(
                sourceCtx, sourceRootFolder, destCtx, destRootFolder,
                destList, "", destFields,
                sourceList.Title, destListName, result);
        }
        else
        {
            Console.WriteLine("[4/5] Migrating items...\n");
            await _itemMigrator.MigrateAllItems(
                sourceCtx, sourceList, destCtx, destList, destFields,
                sourceList.Title, destListName, result);

            if (_config.Options.PreserveAttachments)
            {
                Console.WriteLine("  Migrating attachments...");
                try
                {
                    ListItemCollectionPosition? attachPosition = null;
                    while (true)
                    {
                        var query = new CamlQuery
                        {
                            ListItemCollectionPosition = attachPosition,
                            ViewXml = @"<View Scope='RecursiveAll'><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit>100</RowLimit></View>"
                        };
                        var sourceItems = sourceList.GetItems(query);
                        sourceCtx.Load(sourceItems);
                        sourceCtx.ExecuteQuery();
                        if (sourceItems.Count == 0) break;
                        attachPosition = sourceItems.ListItemCollectionPosition;

                        foreach (var sourceItem in sourceItems)
                        {
                            try
                            {
                                var destId = _manifest.GetDestinationId(sourceList.Title, sourceItem.Id);
                                if (destId == null) continue;

                                var destItem = destList.GetItemById(destId.Value);

                                await _attachmentMigrator.MigrateAttachments(
                                    sourceCtx, sourceItem, destCtx, destItem,
                                    sourceList.Title, result,
                                    destList.RootFolder.ServerRelativeUrl);
                            }
                            catch (Exception attEx)
                            {
                                Console.WriteLine($"  [WARN] Attachment migration failed for item {sourceItem.Id}: {attEx.Message}");
                            }
                        }
                        if (attachPosition == null) break;
                    }
                }
                catch (Exception attEx)
                {
                    Console.WriteLine($"  [WARN] Attachment migration phase failed: {attEx.Message}");
                }
            }
        }

        if (!result.ListStats.ContainsKey(listName))
            result.ListStats[listName] = new ListStats();

        Console.WriteLine($"\n  Completed: '{listName}'");
    }

    private List<string> ResolveListOrder()
    {
        var sourceCtx = _authService.GetContext(_config.SourceSiteUrl);
        sourceCtx.ExecuteQuery();

        var requestedLists = new List<string>();

        if (_config.Lists != null && _config.Lists.Count > 0)
        {
            requestedLists.AddRange(_config.Lists);
        }
        else if (!string.IsNullOrEmpty(_config.SourceListName))
        {
            requestedLists.Add(_config.SourceListName);
        }
        else
        {
            throw new InvalidOperationException(
                "Specify --SourceListName for single list or --Lists for multi-list migration.");
        }

        Console.WriteLine("Analyzing list dependencies...");
        List<string> ordered;
        try
        {
            ordered = _dependencyResolver.ResolveMigrationOrder(
                sourceCtx, requestedLists, _config.AutoResolveDependencies);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [WARN] Dependency analysis failed: {ex.Message}. Falling back to requested order.");
            ordered = requestedLists.ToList();
        }

        Console.WriteLine($"Migration order: {string.Join(" -> ", ordered)}\n");
        return ordered;
    }

    private string GetDestListName(string sourceListName)
    {
        if (_config.Lists != null && _config.Lists.Count > 0)
        {
            if (!string.IsNullOrEmpty(_config.DestinationListName) && _config.Lists.Count == 1)
                return _config.DestinationListName;
            return sourceListName;
        }

        return _config.DestinationListName ?? sourceListName;
    }
}
