using Microsoft.SharePoint.Client;
using MigrationTool2.Helpers;
using MigrationTool2.Models;

namespace MigrationTool2.Services;

public static class ScanService
{
    private static readonly HashSet<FieldType> DirectCopyableTypes = new()
    {
        FieldType.Text, FieldType.Number, FieldType.DateTime, FieldType.Boolean,
        FieldType.Choice, FieldType.MultiChoice, FieldType.Note,
        FieldType.Currency, FieldType.URL, FieldType.Guid,
        FieldType.Integer, FieldType.Counter
    };

    private class FieldInfo
    {
        public string InternalName { get; set; } = "";
        public string Title { get; set; } = "";
        public FieldType TypeKind { get; set; }
        public string TypeName { get; set; } = "";
        public string Category { get; set; } = "";
        public string Detail { get; set; } = "";
        public bool IsBlocker { get; set; }
    }

    public static Task ScanSourceList(MigrationConfig cfg)
    {
        var ctxService = new AuthService(cfg.Auth);
        var srcCtx = ctxService.GetContext(cfg.SourceSiteUrl);

        try { srcCtx.ExecuteQuery(); }
        catch { Console.WriteLine("Failed to connect to source site."); return Task.CompletedTask; }
        Console.WriteLine("Connected.\n");

        var listsToScan = new List<string>();
        if (cfg.Lists != null && cfg.Lists.Count > 0)
            listsToScan.AddRange(cfg.Lists);
        else if (!string.IsNullOrEmpty(cfg.SourceListName))
            listsToScan.Add(cfg.SourceListName);

        if (listsToScan.Count == 0)
        {
            Console.WriteLine("No lists specified. Set SourceListName or Lists in config.");
            return Task.CompletedTask;
        }

        foreach (var listName in listsToScan)
        {
            List? sourceList = null;
            try
            {
                sourceList = srcCtx.Web.Lists.GetByTitle(listName);
                srcCtx.Load(sourceList,
                    l => l.Title, l => l.ItemCount, l => l.BaseTemplate,
                    l => l.BaseType, l => l.ContentTypesEnabled,
                    l => l.Hidden, l => l.Description, l => l.Fields);
                srcCtx.ExecuteQuery();
            }
            catch
            {
                Console.WriteLine($"  List '{listName}' not found on source.\n");
                continue;
            }

            ScanList(sourceList, srcCtx, cfg);
        }

        return Task.CompletedTask;
    }

    private static void ScanList(List list, ClientContext ctx, MigrationConfig cfg)
    {
        var listName = list.Title;
        var separator = new string('═', 50);

        Console.WriteLine($"SCAN REPORT: {listName}");
        Console.WriteLine(separator);
        Console.WriteLine($"  List:         {list.Title}");
        Console.WriteLine($"  Template:     {(int)list.BaseTemplate} ({GetTemplateName((int)list.BaseTemplate)})");
        Console.WriteLine($"  Base Type:    {list.BaseType}");
        Console.WriteLine($"  Items:        {list.ItemCount}");
        Console.WriteLine($"  Hidden:       {list.Hidden}");
        Console.WriteLine($"  Content Type: {(list.ContentTypesEnabled ? "Enabled" : "Disabled")}");
        Console.WriteLine();

        var fields = new List<FieldInfo>();
        var catCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["Simple"] = 0, ["Lookup"] = 0, ["User"] = 0, ["Taxonomy"] = 0, ["Unsupported"] = 0
        };
        var risks = new List<string>();
        var blockers = false;

        foreach (var field in list.Fields)
        {
            if (field.Hidden || field.ReadOnlyField) continue;
            if (field.FieldTypeKind == FieldType.Computed) continue;
            if (field.InternalName.StartsWith("_")) continue;
            if (field.FieldTypeKind == FieldType.Attachments) continue;
            if (field.FieldTypeKind == FieldType.ContentTypeId) continue;

            var fi = new FieldInfo
            {
                InternalName = field.InternalName,
                Title = field.Title,
                TypeKind = field.FieldTypeKind,
                TypeName = field.FieldTypeKind.ToString()
            };

            if (field.FieldTypeKind == FieldType.Lookup)
            {
                try
                {
                    var lookupField = ctx.CastTo<FieldLookup>(field);
                    ctx.Load(lookupField, f => f.LookupList, f => f.LookupField, f => f.AllowMultipleValues);
                    ctx.ExecuteQuery();

                    fi.Category = "Lookup";
                    if (Guid.TryParse(lookupField.LookupList, out var lookupListGuid))
                    {
                        try
                        {
                            var targetList = ctx.Web.Lists.GetById(lookupListGuid);
                            ctx.Load(targetList, l => l.Title);
                            ctx.ExecuteQuery();

                            fi.Detail = $"→ '{targetList.Title}'";
                            if (lookupField.AllowMultipleValues)
                                fi.Detail += " (multi-value)";

                            // Check if target list is in scope
                            var inScope = (cfg.Lists != null && cfg.Lists.Count > 0)
                                ? cfg.Lists.Any(l => string.Equals(l, targetList.Title, StringComparison.OrdinalIgnoreCase))
                                : false;
                            if (!inScope && (cfg.Lists == null || cfg.Lists.Count > 0))
                            {
                                fi.Detail += " ⚠ out of scope";
                                risks.Add($"  ⚠ Lookup '{fi.Title}' targets '{targetList.Title}' which is not in migration scope");
                            }
                        }
                        catch
                        {
                            fi.Detail = "→ (unresolvable) ⚠";
                            risks.Add($"  ⚠ Lookup '{fi.Title}' target list could not be resolved (may be deleted)");
                        }
                    }
                }
                catch
                {
                    fi.Category = "Lookup";
                    fi.Detail = "(could not load)";
                }
            }
            else if (DirectCopyableTypes.Contains(field.FieldTypeKind))
            {
                fi.Category = "Simple";
            }
            else if (field.FieldTypeKind == FieldType.User)
            {
                fi.Category = "User";
                fi.Detail = "needs user resolution";
            }
            else if (FieldHelper.IsTaxonomyField(field))
            {
                fi.Category = "Taxonomy";
                fi.Detail = "needs term store match";
            }
            else
            {
                fi.Category = "Unsupported";
                fi.IsBlocker = true;
                blockers = true;
                fi.Detail = $"cannot migrate this field type";
            }

            fields.Add(fi);
            catCounts.TryGetValue(fi.Category, out var count);
            catCounts[fi.Category] = count + 1;
        }

        Console.WriteLine("  FIELDS");
        Console.WriteLine("  " + new string('─', 85));
        Console.WriteLine($"  {"InternalName",-25} {"Title",-25} {"Type",-15} {"",-2} {"Category",-12} {"Detail",-30}");
        Console.WriteLine("  " + new string('─', 85));

        foreach (var f in fields)
        {
            var icon = f.IsBlocker ? "⛔" : f.Category == "Simple" ? "✓" : "⚠";
            Console.WriteLine($"  {f.InternalName,-25} {f.Title,-25} {f.TypeName,-15} {icon,-2} {f.Category,-12} {f.Detail,-30}");
        }

        Console.WriteLine();
        Console.WriteLine("  CATEGORIES");
        Console.WriteLine($"    Simple:      {catCounts["Simple"]}");
        Console.WriteLine($"    Lookup:      {catCounts["Lookup"]}");
        Console.WriteLine($"    User:        {catCounts["User"]}");
        Console.WriteLine($"    Taxonomy:    {catCounts["Taxonomy"]}");
        Console.WriteLine($"    Unsupported: {catCounts["Unsupported"]}");
        Console.WriteLine();

        // Additional risks
        if (list.ItemCount > 5000)
            risks.Add($"  ⚠ Large list ({list.ItemCount} items) — may hit SharePoint throttling");

        var totalUserFields = fields.Count;
        if (totalUserFields > 20)
            risks.Add($"  ⚠ Many columns ({totalUserFields}) — ValidateUpdateListItem may be slow");

        if (list.BaseTemplate == (int)ListTemplateType.DocumentLibrary)
            risks.Add($"  ⚠ Document library — files handled by FileMigrationService separately");

        if (catCounts["Taxonomy"] > 0)
            risks.Add($"  ⚠ Taxonomy fields require matching terms in destination term store");

        if (catCounts["User"] > 0 && !cfg.Options.ResolveUserDependencies)
            risks.Add($"  ⚠ User fields present but ResolveUserDependencies is disabled");

        if (catCounts["Unsupported"] > 0)
            risks.Add($"  ⛔ Unsupported field types present — migration will fail for those fields");

        if (blockers)
            Console.WriteLine("  BLOCKERS");
        if (risks.Count > 0)
            Console.WriteLine("  RISKS");
        foreach (var risk in risks)
            Console.WriteLine(risk);
        if (blockers || risks.Count > 0)
            Console.WriteLine();

        // Time estimation
        var itemCount = list.ItemCount;
        var complexCount = catCounts["Lookup"] + catCounts["User"] + catCounts["Taxonomy"];
        var baseSec = itemCount * 0.4;
        var complexSec = complexCount * itemCount * 0.1;
        var totalSec = baseSec + complexSec;

        Console.WriteLine($"  ESTIMATED TIME:  {FormatTime(totalSec)}");
        Console.WriteLine($"    Base ({itemCount} items × 0.4s) + Complex ({complexCount} fields × {itemCount} items × 0.1s)");
        Console.WriteLine();

        var complexity = catCounts["Unsupported"] > 0 ? "BLOCKED"
            : catCounts["Lookup"] + catCounts["User"] + catCounts["Taxonomy"] > 5 ? "COMPLEX"
            : catCounts["Lookup"] + catCounts["User"] + catCounts["Taxonomy"] > 0 ? "MODERATE"
            : "SIMPLE";
        Console.WriteLine($"  COMPLEXITY: {complexity}");
        Console.WriteLine(new string('═', 50));
        Console.WriteLine();
    }

    private static string GetTemplateName(int templateId)
    {
        return templateId switch
        {
            100 => "Generic List",
            101 => "Document Library",
            102 => "Survey",
            103 => "Links",
            104 => "Announcements",
            105 => "Contacts",
            106 => "Calendar",
            107 => "Tasks",
            108 => "Discussion Board",
            109 => "Picture Library",
            110 => "Data Sources",
            111 => "Site Template Gallery",
            113 => "Form Library",
            114 => "Site Pages Library",
            115 => "List Template Gallery",
            120 => "Custom Grid in a List",
            170 => "Promoted Links",
            171 => "Tasks with Timeline",
            544 => "Micro Feed",
            700 => "Issue Tracking",
            851 => "Post to List",
            _ => $"Custom ({templateId})"
        };
    }

    private static string FormatTime(double seconds)
    {
        if (seconds < 60)
            return $"~{Math.Max(1, (int)seconds)}s";
        var mins = (int)(seconds / 60);
        var secs = (int)(seconds % 60);
        return $"~{mins}m {secs}s";
    }
}
