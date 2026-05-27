using Microsoft.SharePoint.Client;
using MigrationTool2.Models;

namespace MigrationTool2.Services;

public class ContentTypeMigrationService
{
    private readonly MigrationAuditLog _auditLog;

    public ContentTypeMigrationService(MigrationAuditLog auditLog)
    {
        _auditLog = auditLog;
    }

    public void MigrateContentTypes(ClientContext sourceCtx, List sourceList, ClientContext destCtx, List destList)
    {
        if (!sourceList.ContentTypesEnabled) return;

        try
        {
            sourceCtx.Load(sourceList, l => l.ContentTypes);
            sourceCtx.ExecuteQuery();

            destCtx.Load(destList, l => l.ContentTypes);
            destCtx.ExecuteQuery();

            var destCtNames = destList.ContentTypes
                .AsEnumerable()
                .Select(ct => ct.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var sourceCt in sourceList.ContentTypes)
            {
                if (sourceCt.Name == "Item" || sourceCt.Name == "Document")
                    continue;

                if (destCtNames.Contains(sourceCt.Name))
                {
                    _auditLog.Record(AuditAction.ContentTypeSkipped, sourceList.Title, destList.Title,
                        srcName: sourceCt.Name, destName: sourceCt.Name,
                        status: "Skipped", detail: "Already exists");
                    continue;
                }

                try
                {
                    sourceCtx.Load(sourceCt, ct => ct.Id, ct => ct.StringId, ct => ct.Description, ct => ct.Group);
                    sourceCtx.ExecuteQuery();

                    var ctInfo = new ContentTypeCreationInformation
                    {
                        Name = sourceCt.Name,
                        Description = sourceCt.Description,
                        Group = sourceCt.Group,
                        Id = sourceCt.StringId
                    };

                    var newCt = destCtx.Web.ContentTypes.Add(ctInfo);
                    destCtx.Load(newCt);
                    destCtx.ExecuteQuery();

                    if (newCt != null)
                    {
                        var existingListCt = destList.ContentTypes.FirstOrDefault(ct =>
                            string.Equals(ct.Name, sourceCt.Name, StringComparison.OrdinalIgnoreCase));
                        if (existingListCt == null)
                        {
                            destList.ContentTypes.AddExistingContentType(newCt);
                            destCtx.ExecuteQuery();
                        }
                    }

                    _auditLog.Record(AuditAction.ContentTypeProvisioned, sourceList.Title, destList.Title,
                        srcName: sourceCt.Name, destName: sourceCt.Name);

                    Console.WriteLine($"    + Provisioned content type '{sourceCt.Name}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    [WARN] Could not create content type '{sourceCt.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [WARN] Content type migration failed: {ex.Message}");
        }
    }
}
