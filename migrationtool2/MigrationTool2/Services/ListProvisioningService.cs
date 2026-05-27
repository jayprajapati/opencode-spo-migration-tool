using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using MigrationTool2.Models;
using MigrationTool2.Helpers;

namespace MigrationTool2.Services;

public class ListProvisioningService
{
    private readonly MigrationAuditLog _auditLog;
    private readonly Models.OptionsConfig _options;

    public class DeferredLookupField
    {
        public string InternalName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string SourceLookupListName { get; set; } = string.Empty;
        public string SourceLookupField { get; set; } = string.Empty;
        public string DestinationLookupListName { get; set; } = string.Empty;
        public bool AllowMultipleValues { get; set; }
        public Guid SourceLookupListId { get; set; }
    }

    private static readonly HashSet<FieldType> DirectCopyableTypes = new()
    {
        FieldType.Text, FieldType.Number, FieldType.DateTime, FieldType.Boolean,
        FieldType.Choice, FieldType.MultiChoice, FieldType.Note,
        FieldType.Currency, FieldType.URL, FieldType.Guid,
        FieldType.Integer, FieldType.Counter
    };

    public ListProvisioningService(MigrationAuditLog auditLog, Models.OptionsConfig options)
    {
        _auditLog = auditLog;
        _options = options;
    }

    public (List destList, List<DeferredLookupField> deferredLookups) GetOrCreateList(
        ClientContext ctx, string listName, List sourceList, ClientContext sourceCtx)
    {
        var web = ctx.Web;
        var existingLists = web.Lists;
        ctx.Load(existingLists, ls => ls.Include(l => l.Title));
        ctx.ExecuteQuery();

        var existing = existingLists.FirstOrDefault(l =>
            string.Equals(l.Title, listName, StringComparison.OrdinalIgnoreCase));

        if (existing != null)
        {
            Console.WriteLine($"  Destination list '{listName}' already exists.");
            _auditLog.Record(AuditAction.ListAlreadyExists, sourceList.Title, listName);

            var existingDestList = web.Lists.GetByTitle(listName);
            ctx.Load(existingDestList,
                l => l.Title,
                l => l.Fields.Include(f => f.InternalName, f => f.Title, f => f.Id, f => f.FieldTypeKind, f => f.Hidden, f => f.ReadOnlyField),
                l => l.ContentTypesEnabled, l => l.ContentTypes,
                l => l.EntityTypeName);
            ctx.ExecuteQuery();

            VerifyListName(sourceCtx, sourceList, existingDestList);
            var deferred = ProvisionMissingFields(sourceCtx, sourceList, ctx, existingDestList);
            return (existingDestList, deferred);
        }

        Console.WriteLine($"  Creating destination list '{listName}'...");

        sourceCtx.Load(sourceList,
            l => l.Title, l => l.BaseTemplate, l => l.BaseType,
            l => l.OnQuickLaunch, l => l.ContentTypesEnabled,
            l => l.Hidden, l => l.Description,
            l => l.Fields, l => l.ContentTypes);
        sourceCtx.ExecuteQuery();

        var creationInfo = new ListCreationInformation
        {
            Title = listName,
            TemplateType = (int)sourceList.BaseTemplate,
            Description = sourceList.Description,
            QuickLaunchOption = sourceList.OnQuickLaunch
                ? QuickLaunchOptions.On : QuickLaunchOptions.Off
        };

        var newList = web.Lists.Add(creationInfo);
        newList.ContentTypesEnabled = sourceList.ContentTypesEnabled;
        newList.Hidden = sourceList.Hidden;
        newList.Update();
        ctx.ExecuteQuery();

        _auditLog.Record(AuditAction.ListCreated, sourceList.Title, listName,
            detail: $"Template={(int)sourceList.BaseTemplate}, Hidden={sourceList.Hidden}");

        Console.WriteLine($"  List created. Provisioning fields...");
        var destList = web.Lists.GetByTitle(listName);
        ctx.Load(destList, l => l.Title,
            l => l.Fields.Include(f => f.InternalName, f => f.Title, f => f.Id, f => f.FieldTypeKind, f => f.Hidden, f => f.ReadOnlyField),
            l => l.ContentTypesEnabled, l => l.ContentTypes, l => l.EntityTypeName);
        ctx.ExecuteQuery();

        VerifyListName(sourceCtx, sourceList, destList);
        var deferredLookups = ProvisionFields(sourceCtx, sourceList, ctx, destList);

        ctx.Load(destList, l => l.Fields);
        ctx.ExecuteQuery();

        Console.WriteLine($"  Provisioned {destList.Fields.Count} fields.");
        return (destList, deferredLookups);
    }

    private void VerifyListName(ClientContext sourceCtx, List sourceList, List destList)
    {
        _auditLog.RecordNameVerification(
            sourceList.Title, destList.Title,
            "List.Title", sourceList.Title, destList.Title,
            string.Equals(sourceList.Title, destList.Title, StringComparison.Ordinal));

        try
        {
            sourceCtx.Load(sourceList, l => l.EntityTypeName);
            sourceCtx.ExecuteQuery();
        }
        catch { }

        if (!string.IsNullOrEmpty(sourceList.EntityTypeName))
        {
            _auditLog.RecordNameVerification(
                sourceList.Title, destList.Title,
                "List.EntityTypeName", sourceList.EntityTypeName, destList.EntityTypeName,
                string.Equals(sourceList.EntityTypeName, destList.EntityTypeName, StringComparison.Ordinal));
        }
    }

    private void VerifyFieldNames(List sourceList, string destListName, Field sourceField, Field destField)
    {
        var sourceInternal = sourceField.InternalName;
        var destInternal = destField.InternalName;
        var internalMatch = string.Equals(sourceInternal, destInternal, StringComparison.Ordinal);

        _auditLog.RecordNameVerification(
            sourceList.Title, destListName,
            "Field.InternalName", sourceInternal, destInternal,
            internalMatch,
            internalMatch ? null : $"Source field '{sourceField.Title}' internal='{sourceInternal}' dest='{destInternal}'");

        var sourceDisplay = sourceField.Title;
        var destDisplay = destField.Title;
        var displayMatch = string.Equals(sourceDisplay, destDisplay, StringComparison.Ordinal);

        _auditLog.RecordNameVerification(
            sourceList.Title, destListName,
            "Field.Title", sourceDisplay, destDisplay,
            displayMatch,
            displayMatch ? null : $"InternalName='{sourceInternal}' source display='{sourceDisplay}' dest='{destDisplay}'");
    }

    private List<DeferredLookupField> ProvisionFields(
        ClientContext sourceCtx, List sourceList,
        ClientContext ctx, List destList)
    {
        var deferred = new List<DeferredLookupField>();
        var destInternalNames = destList.Fields.AsEnumerable()
            .Select(f => f.InternalName).ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var sourceField in sourceList.Fields)
        {
            if (sourceField.Hidden || sourceField.ReadOnlyField)
                continue;
            if (sourceField.FieldTypeKind == FieldType.Computed)
                continue;
            if (sourceField.InternalName.StartsWith("_"))
                continue;

            var internalName = sourceField.InternalName;

            if (sourceField.FieldTypeKind == FieldType.Lookup)
            {
                var lookupField = sourceCtx.CastTo<FieldLookup>(sourceField);
                sourceCtx.Load(lookupField,
                    f => f.LookupList, f => f.LookupField,
                    f => f.Title, f => f.InternalName,
                    f => f.AllowMultipleValues);
                sourceCtx.ExecuteQuery();

                if (!Guid.TryParse(lookupField.LookupList, out var lookupListGuid)) continue;

                try
                {
                    var lookupList = sourceCtx.Web.Lists.GetById(lookupListGuid);
                    sourceCtx.Load(lookupList, l => l.Title);
                    sourceCtx.ExecuteQuery();

                    deferred.Add(new DeferredLookupField
                    {
                        InternalName = lookupField.InternalName,
                        Title = lookupField.Title,
                        SourceLookupListName = lookupList.Title,
                        SourceLookupField = lookupField.LookupField,
                        AllowMultipleValues = lookupField.AllowMultipleValues,
                        SourceLookupListId = lookupListGuid
                    });

                    _auditLog.Record(AuditAction.LookupFieldDeferred, sourceList.Title, destList.Title,
                        srcId: internalName,
                        srcName: sourceField.Title,
                        detail: $"Depends on list '{lookupList.Title}'");
                }
                catch
                {
                    Console.WriteLine($"    [WARN] Could not resolve lookup target for field '{lookupField.Title}'. Skipping.");
                }

                continue;
            }

            if (destInternalNames.Contains(internalName))
            {
                var destField = destList.Fields.FirstOrDefault(f =>
                    string.Equals(f.InternalName, internalName, StringComparison.OrdinalIgnoreCase));
                if (destField != null)
                {
                    VerifyFieldNames(sourceList, destList.Title, sourceField, destField);
                    _auditLog.Record(AuditAction.FieldSkipped, sourceList.Title, destList.Title,
                        srcId: sourceField.InternalName, destId: destField.InternalName,
                        srcName: sourceField.Title, destName: destField.Title,
                        status: "Skipped", detail: "Already exists");
                }
                continue;
            }

            if (sourceField.FieldTypeKind == FieldType.User)
            {
                if (_options.ResolveUserDependencies)
                {
                    TryAddUserField(ctx, destList, sourceField, internalName);
                    destInternalNames.Add(internalName);
                }
                else
                {
                    Console.WriteLine($"    Skipping field '{sourceField.Title}' ({internalName}): User type (disabled)");
                }
                continue;
            }

            if (FieldHelper.IsTaxonomyField(sourceField))
            {
                TryAddTaxonomyField(ctx, destList, sourceField, internalName);
                destInternalNames.Add(internalName);
                continue;
            }

            TryAddField(sourceCtx, sourceList, ctx, destList, sourceField, internalName);
            destInternalNames.Add(internalName);
        }

        return deferred;
    }

    private List<DeferredLookupField> ProvisionMissingFields(
        ClientContext sourceCtx, List sourceList,
        ClientContext ctx, List destList)
    {
        ctx.Load(destList,
            l => l.Fields.Include(f => f.InternalName, f => f.Title, f => f.Id, f => f.FieldTypeKind, f => f.Hidden, f => f.ReadOnlyField));
        ctx.ExecuteQuery();

        var deferred = new List<DeferredLookupField>();
        var destInternalNames = destList.Fields.AsEnumerable()
            .Select(f => f.InternalName).ToHashSet(StringComparer.OrdinalIgnoreCase);

        sourceCtx.Load(sourceList, l => l.Fields);
        sourceCtx.ExecuteQuery();

        foreach (var sourceField in sourceList.Fields)
        {
            if (sourceField.Hidden || sourceField.ReadOnlyField)
                continue;
            if (sourceField.FieldTypeKind == FieldType.Computed)
                continue;
            if (sourceField.InternalName.StartsWith("_"))
                continue;

            var internalName = sourceField.InternalName;

            if (sourceField.FieldTypeKind == FieldType.Lookup)
            {
                var lookupField = sourceCtx.CastTo<FieldLookup>(sourceField);
                sourceCtx.Load(lookupField,
                    f => f.LookupList, f => f.LookupField,
                    f => f.Title, f => f.InternalName,
                    f => f.AllowMultipleValues);
                sourceCtx.ExecuteQuery();

                if (!Guid.TryParse(lookupField.LookupList, out var lookupListGuid)) continue;

                try
                {
                    var lookupList = sourceCtx.Web.Lists.GetById(lookupListGuid);
                    sourceCtx.Load(lookupList, l => l.Title);
                    sourceCtx.ExecuteQuery();

                    deferred.Add(new DeferredLookupField
                    {
                        InternalName = lookupField.InternalName,
                        Title = lookupField.Title,
                        SourceLookupListName = lookupList.Title,
                        SourceLookupField = lookupField.LookupField,
                        AllowMultipleValues = lookupField.AllowMultipleValues,
                        SourceLookupListId = lookupListGuid
                    });

                    _auditLog.Record(AuditAction.LookupFieldDeferred, sourceList.Title, destList.Title,
                        srcId: internalName,
                        srcName: sourceField.Title,
                        detail: $"Depends on list '{lookupList.Title}'");
                }
                catch
                {
                    Console.WriteLine($"    [WARN] Could not resolve lookup target for field '{lookupField.Title}'. Skipping.");
                }

                continue;
            }

            if (destInternalNames.Contains(internalName))
            {
                var destField = destList.Fields.FirstOrDefault(f =>
                    string.Equals(f.InternalName, internalName, StringComparison.OrdinalIgnoreCase));
                if (destField != null)
                {
                    VerifyFieldNames(sourceList, destList.Title, sourceField, destField);
                }
                continue;
            }

            if (sourceField.FieldTypeKind == FieldType.User)
            {
                if (_options.ResolveUserDependencies)
                {
                    TryAddUserField(ctx, destList, sourceField, internalName);
                    destInternalNames.Add(internalName);
                }
                continue;
            }

            if (FieldHelper.IsTaxonomyField(sourceField))
            {
                TryAddTaxonomyField(ctx, destList, sourceField, internalName);
                destInternalNames.Add(internalName);
                continue;
            }

            TryAddField(sourceCtx, sourceList, ctx, destList, sourceField, internalName);
            destInternalNames.Add(internalName);
        }

        ctx.Load(destList,
            l => l.Fields.Include(f => f.InternalName, f => f.Title, f => f.Id, f => f.FieldTypeKind, f => f.Hidden, f => f.ReadOnlyField));
        ctx.ExecuteQuery();

        EnsureFieldsOnDefaultContentType(ctx, destList, sourceList, sourceCtx);
        return deferred;
    }

    private void EnsureFieldsOnDefaultContentType(ClientContext ctx, List destList, List sourceList, ClientContext sourceCtx)
    {
        try
        {
            sourceCtx.Load(sourceList, l => l.Fields);
            sourceCtx.ExecuteQuery();

            ctx.Load(destList, l => l.ContentTypes);
            ctx.ExecuteQuery();

            var defaultCt = destList.ContentTypes.FirstOrDefault();
            if (defaultCt == null) return;

            ctx.Load(defaultCt, ct => ct.FieldLinks);
            ctx.ExecuteQuery();

            var existingFieldLinkNames = defaultCt.FieldLinks
                .AsEnumerable()
                .Select(fl => fl.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var sourceField in sourceList.Fields)
            {
                if (sourceField.Hidden || sourceField.ReadOnlyField) continue;
                if (sourceField.FieldTypeKind == FieldType.Computed) continue;
                if (sourceField.InternalName.StartsWith("_")) continue;

                var internalName = sourceField.InternalName;
                if (existingFieldLinkNames.Contains(internalName)) continue;

                var destField = destList.Fields.FirstOrDefault(f =>
                    string.Equals(f.InternalName, internalName, StringComparison.OrdinalIgnoreCase));
                if (destField == null) continue;

                var flci = new FieldLinkCreationInformation { Field = destField };
                defaultCt.FieldLinks.Add(flci);
                defaultCt.Update(true);
                ctx.ExecuteQuery();
                existingFieldLinkNames.Add(internalName);

                Console.WriteLine($"    + Added field '{sourceField.Title}' ({internalName}) to default content type");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    [WARN] Could not ensure fields on default content type: {ex.Message}");
        }
    }

    private void TryAddField(ClientContext sourceCtx, List sourceList, ClientContext ctx, List destList, Field sourceField, string internalName)
    {
        try
        {
            if (!DirectCopyableTypes.Contains(sourceField.FieldTypeKind))
            {
                Console.WriteLine($"    Skipping field '{sourceField.Title}' ({internalName}): unsupported type {sourceField.FieldTypeKind}");
                return;
            }

            var fieldXml = sourceField.SchemaXml;
            fieldXml = fieldXml.Replace("Group=\"\"", "Group=\"MigrationTool\"");
            fieldXml = fieldXml.Replace($"ID=\"{sourceField.Id}\"", $"ID=\"{Guid.NewGuid()}\"");

            destList.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToDefaultContentType);
            ctx.ExecuteQuery();

            Console.WriteLine($"    + Provisioned field '{sourceField.Title}' ({internalName})");

            _auditLog.Record(AuditAction.FieldProvisioned, sourceList.Title, destList.Title,
                srcId: internalName, destId: internalName,
                srcName: sourceField.Title, detail: $"Type={sourceField.FieldTypeKind}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    ! FAILED to provision field '{sourceField.Title}' ({internalName}): {ex.GetType().Name}: {ex.Message}");
            _auditLog.Record(AuditAction.FieldProvisionFailed, sourceList.Title, destList.Title,
                srcId: internalName,
                srcName: sourceField.Title,
                status: "Failed",
                detail: ex.Message);
        }
    }

    private void TryAddUserField(ClientContext ctx, List destList, Field sourceField, string internalName)
    {
        try
        {
            var fieldXml = sourceField.SchemaXml;
            fieldXml = fieldXml.Replace("Group=\"\"", "Group=\"MigrationTool\"");
            fieldXml = fieldXml.Replace($"ID=\"{sourceField.Id}\"", $"ID=\"{Guid.NewGuid()}\"");

            destList.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToDefaultContentType);
            ctx.ExecuteQuery();

            Console.WriteLine($"    + Provisioned User field '{sourceField.Title}' ({internalName})");

            _auditLog.Record(AuditAction.FieldProvisioned, destList.Title, destList.Title,
                srcId: internalName, destId: internalName,
                srcName: sourceField.Title, detail: $"Type=User");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    ! FAILED to provision User field '{sourceField.Title}' ({internalName}): {ex.Message}");
            _auditLog.Record(AuditAction.FieldProvisionFailed, destList.Title, destList.Title,
                srcId: internalName,
                srcName: sourceField.Title,
                status: "Failed",
                detail: ex.Message);
        }
    }

    private void TryAddTaxonomyField(ClientContext ctx, List destList, Field sourceField, string internalName)
    {
        try
        {
            var fieldXml = sourceField.SchemaXml;
            fieldXml = fieldXml.Replace("Group=\"\"", "Group=\"MigrationTool\"");
            fieldXml = fieldXml.Replace($"ID=\"{sourceField.Id}\"", $"ID=\"{Guid.NewGuid()}\"");

            destList.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToDefaultContentType);
            ctx.ExecuteQuery();

            Console.WriteLine($"    + Provisioned Taxonomy field '{sourceField.Title}' ({internalName})");

            _auditLog.Record(AuditAction.FieldProvisioned, destList.Title, destList.Title,
                srcId: internalName, destId: internalName,
                srcName: sourceField.Title, detail: $"Type=TaxonomyField");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    ! FAILED to provision Taxonomy field '{sourceField.Title}' ({internalName}): {ex.Message}");
            _auditLog.Record(AuditAction.FieldProvisionFailed, destList.Title, destList.Title,
                srcId: internalName,
                srcName: sourceField.Title,
                status: "Failed",
                detail: ex.Message);
        }
    }

    public void CreateLookupFieldOnDestination(
        ClientContext ctx, List destList,
        DeferredLookupField deferred,
        string destinationLookupListName,
        ClientContext sourceCtx)
    {
        try
        {
            var targetList = ctx.Web.Lists.GetByTitle(destinationLookupListName);
            ctx.Load(targetList, l => l.Id);
            ctx.ExecuteQuery();

            var fieldSchema = $@"<Field Type='Lookup' DisplayName='{deferred.Title}' Name='{deferred.InternalName}' 
                StaticName='{deferred.InternalName}' Hidden='FALSE' Required='FALSE' 
                List='{{{targetList.Id}}}' ShowField='{deferred.SourceLookupField}' 
                UnlimitedLengthInDocumentLibrary='FALSE' Group='MigrationTool'/>";

            destList.Fields.AddFieldAsXml(fieldSchema, false, AddFieldOptions.AddToDefaultContentType);
            ctx.ExecuteQuery();

            Console.WriteLine($"    + Created Lookup field '{deferred.Title}' -> '{destinationLookupListName}'");

            _auditLog.Record(AuditAction.LookupFieldCreated, deferred.SourceLookupListName, destList.Title,
                srcId: deferred.InternalName,
                srcName: deferred.Title,
                detail: $"Target list='{destinationLookupListName}', ShowField='{deferred.SourceLookupField}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"    ! FAILED to create Lookup field '{deferred.Title}': {ex.Message}");
            _auditLog.Record(AuditAction.FieldProvisionFailed, deferred.SourceLookupListName, destList.Title,
                srcId: deferred.InternalName,
                srcName: deferred.Title,
                status: "Failed",
                detail: ex.Message);
        }
    }

    public Dictionary<string, Field> GetFieldMap(List list, ClientContext ctx)
    {
        ctx.Load(list, l => l.Fields);
        ctx.ExecuteQuery();

        var allowedHidden = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "Title", "Author", "Editor", "Created", "Modified" };

        return list.Fields
            .AsEnumerable()
            .Where(f => !f.Hidden || allowedHidden.Contains(f.InternalName))
            .Where(f => !f.ReadOnlyField)
            .Where(f => f.FieldTypeKind != FieldType.Computed)
            .Where(f => !f.InternalName.StartsWith("_"))
            .ToDictionary(f => f.InternalName, StringComparer.OrdinalIgnoreCase);
    }
}
