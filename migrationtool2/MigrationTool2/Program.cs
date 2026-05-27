using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.SharePoint.Client;
using MigrationTool2.Helpers;
using MigrationTool2.Models;
using MigrationTool2.Services;

var assemblyDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;
var builder = new ConfigurationBuilder()
    .SetBasePath(assemblyDir)
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddCommandLine(args);

var configuration = builder.Build();

var config = configuration.Get<MigrationConfig>()
    ?? throw new InvalidOperationException("Failed to load configuration");

var services = new ServiceCollection();

services.AddSingleton(config.Auth);
services.AddSingleton(config.Options);
services.AddSingleton(config);

services.AddSingleton<AuthService>();
services.AddSingleton<ListProvisioningService>();
services.AddSingleton<DependencyResolverService>();
services.AddSingleton<ContentTypeMigrationService>();
services.AddSingleton<ThrottleHandler>(_ => new ThrottleHandler(config.Options.MaxRetryCount));

var manifestName = BuildManifestFileName(config, assemblyDir);
var manifest = new MigrationManifest(Path.Combine(assemblyDir, manifestName));
manifest.LoadExisting();
services.AddSingleton(manifest);

var auditLog = new MigrationAuditLog(Path.Combine(assemblyDir, "migration_audit.csv"));
services.AddSingleton(auditLog);

var idMapping = manifest.GetAllMappings();
var idMappingWritable = idMapping.ToDictionary(
    kv => kv.Key,
    kv => kv.Value.ToDictionary(kv2 => kv2.Key, kv2 => kv2.Value));
services.AddSingleton<AuthService>(sp => new AuthService(config.Auth));
services.AddSingleton(sp =>
{
    var authService = sp.GetRequiredService<AuthService>();
    return new FieldHelper(config.Options, idMappingWritable, authService, config.DestinationSiteUrl);
});

services.AddSingleton<FileMigrationService>();
services.AddSingleton<ItemMigrationService>();
services.AddSingleton<AttachmentMigrationService>();
services.AddSingleton<MigrationOrchestrator>();

var serviceProvider = services.BuildServiceProvider();

if (args.Any(a => string.Equals(a, "--verify", StringComparison.OrdinalIgnoreCase)))
{
    await VerifyPreserveFields(config);
    return;
}

if (args.Any(a => string.Equals(a, "--clean", StringComparison.OrdinalIgnoreCase)))
{
    await CleanDestinationList(config);
    return;
}

if (args.Any(a => string.Equals(a, "--test-rest", StringComparison.OrdinalIgnoreCase)))
{
    await TestRestApiWrite(config);
    return;
}

if (args.Any(a => string.Equals(a, "--scan", StringComparison.OrdinalIgnoreCase)))
{
    await ScanService.ScanSourceList(config);
    return;
}

Console.WriteLine("SharePoint Migration Tool v2");
Console.WriteLine("==========================\n");

var orchestrator = serviceProvider.GetRequiredService<MigrationOrchestrator>();
await orchestrator.ExecuteAsync();

static string BuildManifestFileName(MigrationConfig cfg, string assemblyDir)
{
    if (cfg.Lists != null && cfg.Lists.Count > 0)
    {
        var listPart = string.Join("_", cfg.Lists).Replace(" ", "_");
        return $"migration_manifest_{listPart}.csv";
    }
    if (!string.IsNullOrEmpty(cfg.SourceListName))
    {
        var dest = cfg.DestinationListName ?? cfg.SourceListName;
        return $"migration_manifest_{cfg.SourceListName}_{dest}.csv";
    }
    return "migration_manifest.csv";
}

static async Task VerifyPreserveFields(MigrationConfig cfg)
{
    var ctxService = new AuthService(cfg.Auth);
    var destCtx = ctxService.GetContext(cfg.DestinationSiteUrl);
    var srcCtx = ctxService.GetContext(cfg.SourceSiteUrl);

    try { destCtx.ExecuteQuery(); } catch { Console.WriteLine("Destination: FAILED"); return; }
    try { srcCtx.ExecuteQuery(); } catch { Console.WriteLine("Source: FAILED"); return; }
    Console.WriteLine("Connected.");

    var listsToVerify = new List<string>();
    if (cfg.Lists != null && cfg.Lists.Count > 0)
        listsToVerify.AddRange(cfg.Lists);
    else if (!string.IsNullOrEmpty(cfg.SourceListName))
        listsToVerify.Add(cfg.SourceListName);

    var token = await ctxService.GetAccessTokenAsync(cfg.DestinationSiteUrl);
    using var restClient = new HttpClient();
    restClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
    restClient.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");

    foreach (var listName in listsToVerify)
    {
        Console.WriteLine($"\nVerifying: '{listName}'");

        List? srcList = null;
        try
        {
            srcList = srcCtx.Web.Lists.GetByTitle(listName);
            srcCtx.Load(srcList);
            srcCtx.ExecuteQuery();
        }
        catch { Console.WriteLine($"  Source list '{listName}' not found."); continue; }

        var destName = cfg.DestinationListName ?? listName;
        var baseUrl = cfg.DestinationSiteUrl.TrimEnd('/');

        // Load source items via CSOM (works with CSOM)
        var viewFields = @"<FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Author'/><FieldRef Name='Editor'/><FieldRef Name='Created'/><FieldRef Name='Modified'/>";
        var srcQuery = new CamlQuery { ViewXml = $"<View Scope='RecursiveAll'><ViewFields>{viewFields}</ViewFields><RowLimit>5000</RowLimit></View>" };
        var srcItems = srcList.GetItems(srcQuery);
        srcCtx.Load(srcItems);
        srcCtx.ExecuteQuery();

        // Load destination items via REST API (CSOM can't read fields with app-only)
        var restUrl = $"{baseUrl}/_api/web/lists/getbytitle('{destName.Replace("'", "''")}')/items?$select=Id,Title,Created,Modified,AuthorId,EditorId&$top=5000";
        var restResponse = await restClient.GetAsync(restUrl);
        var restBody = await restResponse.Content.ReadAsStringAsync();
        if (!restResponse.IsSuccessStatusCode)
        {
            Console.WriteLine($"  REST verify failed: {restBody}");
            continue;
        }
        var doc = JsonDocument.Parse(restBody);
        var destItemsData = doc.RootElement.GetProperty("d").GetProperty("results");

        var destByTitle = new Dictionary<string, Queue<JsonElement>>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in destItemsData.EnumerateArray())
        {
            try
            {
                var t = item.GetProperty("Title").GetString();
                if (!string.IsNullOrEmpty(t))
                {
                    if (!destByTitle.ContainsKey(t))
                        destByTitle[t] = new Queue<JsonElement>();
                    destByTitle[t].Enqueue(item);
                }
            }
            catch (Exception ex) { Console.Error.WriteLine($"  [DBG] Title access failed: {ex.GetType().Name}: {ex.Message}"); }
        }

        Console.WriteLine($"\n{"Title",-15} {"Created (src)",-22} {"Created (dst)",-22} {"Modified (src)",-22} {"Modified (dst)",-22} {"AuthorId src",-12} {"AuthorId dst",-12}");
        Console.WriteLine(new string('-', 120));

        int pass = 0, fail = 0;
        foreach (var src in srcItems)
        {
            string sTitle = src["Title"]?.ToString() ?? "";
            string sCreated = "", sModified = "";
            int sAuthorId = 0;

            var c = src["Created"];
            if (c is DateTime dt) sCreated = dt.ToString("o");
            else if (c is string cs) sCreated = cs;

            var m = src["Modified"];
            if (m is DateTime mdt) sModified = mdt.ToString("o");
            else if (m is string ms) sModified = ms;

            var sa = src["Author"];
            if (sa is FieldUserValue fva) sAuthorId = fva.LookupId;

            if (!destByTitle.TryGetValue(sTitle, out var queue) || queue.Count == 0)
            {
                Console.WriteLine($"{sTitle,-15} MISSING ON DESTINATION");
                fail++;
                continue;
            }

            var dst = queue.Dequeue();

            var dCreated = dst.GetProperty("Created").GetString();
            var dModified = dst.GetProperty("Modified").GetString();
            var dAuthorId = dst.GetProperty("AuthorId").GetInt32();

            // Compare dates with tolerance for seconds truncation (SharePoint rounds to minute)
            var cOk = DatesMatchWithinTolerance(sCreated, dCreated, TimeSpan.FromSeconds(70));
            var mOk = DatesMatchWithinTolerance(sModified, dModified, TimeSpan.FromSeconds(70));
            // AuthorId compare — different sites give different IDs for same user (expected)
            var allOk = cOk && mOk;
            if (allOk) pass++; else fail++;

            Console.WriteLine($"{sTitle,-15} {sCreated,-22} {dCreated,-22} {sModified,-22} {dModified,-22} {sAuthorId,-12} {dAuthorId,-12} {(allOk ? "PASS" : "FAIL")}");
        }
        Console.WriteLine($"\n  {listName}: Pass={pass}, Fail={fail}");
        Console.WriteLine("  (AuthorId differences across sites are expected and not counted as failures)");
    }
}

static bool DatesMatchWithinTolerance(string dateStr1, string dateStr2, TimeSpan tolerance)
{
    if (DateTime.TryParse(dateStr1, null, System.Globalization.DateTimeStyles.RoundtripKind, out var d1) &&
        DateTime.TryParse(dateStr2, null, System.Globalization.DateTimeStyles.RoundtripKind, out var d2))
    {
        var diff = (d1 - d2).Duration();
        return diff <= tolerance;
    }
    return string.Equals(dateStr1, dateStr2, StringComparison.OrdinalIgnoreCase);
}

static async Task TestRestApiWrite(MigrationConfig cfg)
{
    var ctxService = new AuthService(cfg.Auth);
    var token = await ctxService.GetAccessTokenAsync(cfg.DestinationSiteUrl);

    using var client = new HttpClient();
    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
    client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");

    var listName = cfg.DestinationListName ?? cfg.SourceListName ?? "Cars";
    var baseUrl = cfg.DestinationSiteUrl.TrimEnd('/');

    // Create a separate test item
    var url = $"{baseUrl}/_api/web/lists/getbytitle('{listName}')/items";
    var createJson = JsonSerializer.Serialize(new { Title = "Date Format Test", Model = "Test" });
    var createResp = await client.PostAsync(url, new StringContent(createJson, Encoding.UTF8, "application/json"));
    var createBody = await createResp.Content.ReadAsStringAsync();
    var doc = JsonDocument.Parse(createBody);
    var newId = doc.RootElement.GetProperty("d").GetProperty("Id").GetInt32();
    Console.WriteLine($"Created item {newId}");

    // Test Author with different formats
    foreach (var (label, value) in new (string, string)[] {
        ("email", "jay@vkn4v.onmicrosoft.com"),
        ("userId", "13"),
        ("claim", "[{'Key':'i:0#.f|membership|jay@vkn4v.onmicrosoft.com'}]") })
    {
        // Create fresh item per test
        var cr = await client.PostAsync(url, new StringContent(
            JsonSerializer.Serialize(new { Title = $"AuthorTest-{label}", Model = "X" }),
            Encoding.UTF8, "application/json"));
        var cb = await cr.Content.ReadAsStringAsync();
        var cid = JsonDocument.Parse(cb).RootElement.GetProperty("d").GetProperty("Id").GetInt32();

        var vurl = $"{baseUrl}/_api/web/lists/getbytitle('{listName}')/items({cid})/ValidateUpdateListItem";
        var vbody = JsonSerializer.Serialize(new
        {
            formValues = new[] {
                new { FieldName = "Author", FieldValue = value },
                new { FieldName = "Created", FieldValue = "03/26/2024 09:01:09" }
            },
            bNewDocumentUpdate = true
        });
        var vresp = await client.PostAsync(vurl, new StringContent(vbody, Encoding.UTF8, "application/json"));
        var bodyText = await vresp.Content.ReadAsStringAsync();
        Console.WriteLine($"Author format '{label}'='{value}': {bodyText}");

        // Check result
        var gr = await client.GetAsync($"{baseUrl}/_api/web/lists/getbytitle('{listName}')/items({cid})?$select=Id,Title,AuthorId,Created");
        Console.WriteLine($"  => {(await gr.Content.ReadAsStringAsync())}\n");
    }
}

static async Task CleanDestinationList(MigrationConfig cfg)
{
    var ctxService = new AuthService(cfg.Auth);
    var ctx = ctxService.GetContext(cfg.DestinationSiteUrl);
    try { ctx.ExecuteQuery(); } catch { Console.WriteLine("Failed to connect"); return; }

    var listsToClean = new List<string>();
    if (cfg.Lists != null && cfg.Lists.Count > 0)
        listsToClean.AddRange(cfg.Lists);
    else if (!string.IsNullOrEmpty(cfg.DestinationListName))
        listsToClean.Add(cfg.DestinationListName);
    else if (!string.IsNullOrEmpty(cfg.SourceListName))
        listsToClean.Add(cfg.SourceListName);

    foreach (var listName in listsToClean)
    {
        try
        {
            var list = ctx.Web.Lists.GetByTitle(listName);
            list.DeleteObject();
            ctx.ExecuteQuery();
            Console.WriteLine($"Deleted list '{listName}'");
        }
        catch (Exception ex) { Console.WriteLine($"Could not delete '{listName}': {ex.Message}"); }
    }
}
