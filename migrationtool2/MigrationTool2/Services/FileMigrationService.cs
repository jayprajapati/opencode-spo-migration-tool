using Microsoft.SharePoint.Client;
using MigrationTool2.Helpers;
using MigrationTool2.Models;

namespace MigrationTool2.Services;

public class FileMigrationService
{
    private readonly ThrottleHandler _throttle;
    private readonly OptionsConfig _options;
    private readonly MigrationManifest _manifest;
    private readonly FieldHelper _fieldCopier;
    private readonly MigrationAuditLog _auditLog;
    private readonly AttachmentMigrationService _attachmentMigrator;

    public FileMigrationService(
        ThrottleHandler throttle,
        OptionsConfig options,
        MigrationManifest manifest,
        FieldHelper fieldCopier,
        MigrationAuditLog auditLog,
        AttachmentMigrationService attachmentMigrator)
    {
        _throttle = throttle;
        _options = options;
        _manifest = manifest;
        _fieldCopier = fieldCopier;
        _auditLog = auditLog;
        _attachmentMigrator = attachmentMigrator;
    }

    public async Task MigrateFilesInFolder(
        ClientContext sourceCtx, Folder sourceFolder,
        ClientContext destCtx, Folder destFolder,
        List destList, string relativePath,
        Dictionary<string, Field> destFields,
        string sourceListName, string destListName,
        MigrationResult result)
    {
        sourceCtx.Load(sourceFolder, f => f.Files);
        sourceCtx.Load(sourceFolder, f => f.Folders);
        sourceCtx.ExecuteQuery();

        foreach (var file in sourceFolder.Files)
        {
            sourceCtx.Load(file, f => f.ListItemAllFields, f => f.Name, f => f.ServerRelativeUrl, f => f.Length, f => f.TimeLastModified);
            sourceCtx.ExecuteQuery();

            var sourceItem = file.ListItemAllFields;

            var fileTitle = sourceItem.FieldValues.GetValueOrDefault("Title")?.ToString() ?? file.Name;
            var filePath = string.IsNullOrEmpty(relativePath)
                ? file.Name
                : $"{relativePath}/{file.Name}";

            if (_options.OverwriteMode == "Skip" && _manifest.IsAlreadyMigrated(sourceListName, (int)sourceItem.Id))
            {
                result.Skipped++;
                _auditLog.Record(AuditAction.ItemSkipped, sourceListName, destListName,
                    srcId: sourceItem.Id.ToString(),
                    srcName: filePath,
                    status: "Skipped", detail: "OverwriteMode=Skip, already migrated");
                continue;
            }

            try
            {
                await _throttle.ExecuteWithRetryAsync(() =>
                {
                    var existingFile = destFolder.Files.FirstOrDefault(f =>
                        string.Equals(f.Name, file.Name, StringComparison.OrdinalIgnoreCase));

                    if (existingFile != null)
                    {
                        if (_options.OverwriteMode == "Skip")
                        {
                            result.Skipped++;
                            return Task.CompletedTask;
                        }

                        existingFile.DeleteObject();
                        destCtx.ExecuteQuery();
                    }

                    var stream = file.OpenBinaryStream();
                    sourceCtx.ExecuteQuery();

                    var fileCreationInfo = new FileCreationInformation
                    {
                        ContentStream = stream.Value,
                        Url = file.Name,
                        Overwrite = true
                    };

                    var uploadedFile = destFolder.Files.Add(fileCreationInfo);
                    destCtx.Load(uploadedFile, f => f.ListItemAllFields, f => f.Name);
                    destCtx.ExecuteQuery();

                    var destItem = uploadedFile.ListItemAllFields;
                    destCtx.Load(destItem, i => i.Id);
                    destCtx.ExecuteQuery();

                    var destId = destItem.Id;

                    _fieldCopier.CreateAndCopyFields(
                        destList, sourceItem, sourceCtx, destCtx, destFields,
                        sourceListName, destListName, existingItem: destItem);

                    if (_options.PreserveAttachments)
                    {
                        try
                        {
                            _attachmentMigrator.MigrateAttachments(
                                sourceCtx, sourceItem, destCtx, destItem,
                                sourceListName, result,
                                destList.RootFolder.ServerRelativeUrl).Wait();
                        }
                        catch { }
                    }

                    _manifest.RecordItem(sourceListName, (int)sourceItem.Id, destId, fileTitle, "OK");
                    result.Succeeded++;

                    _auditLog.Record(AuditAction.FileUploaded, sourceListName, destListName,
                        srcId: sourceItem.Id.ToString(), destId: destId.ToString(),
                        srcName: filePath, destName: filePath,
                        detail: $"Size={file.Length}");

                    Console.WriteLine($"  + {filePath}");

                    return Task.CompletedTask;
                }, $"File: {filePath}");
            }
            catch (Exception ex)
            {
                result.Failed++;
                result.Errors.Add(new MigrationError
                {
                    ListName = sourceListName,
                    ItemId = sourceItem.Id.ToString(),
                    Title = filePath,
                    Message = ex.Message,
                    Exception = ex
                });

                _auditLog.Record(AuditAction.ItemFailed, sourceListName, destListName,
                    srcId: sourceItem.Id.ToString(),
                    srcName: filePath,
                    status: "Failed", detail: ex.Message);

                Console.WriteLine($"  ! FAILED: {filePath} - {ex.Message}");
            }
        }

        foreach (var subFolder in sourceFolder.Folders)
        {
            if (subFolder.Name is "Forms" or "_cts") continue;

            sourceCtx.Load(subFolder, f => f.Name, f => f.ServerRelativeUrl);
            sourceCtx.ExecuteQuery();

            var subRelativePath = string.IsNullOrEmpty(relativePath)
                ? subFolder.Name
                : $"{relativePath}/{subFolder.Name}";

            var freshSrcSub = sourceCtx.Web.GetFolderByServerRelativeUrl(subFolder.ServerRelativeUrl);
            sourceCtx.Load(freshSrcSub, f => f.ServerRelativeUrl, f => f.Files, f => f.Folders);
            sourceCtx.ExecuteQuery();

            var destSubFolder = destFolder.Folders.FirstOrDefault(f =>
                string.Equals(f.Name, subFolder.Name, StringComparison.OrdinalIgnoreCase));

            if (destSubFolder == null)
            {
                destSubFolder = destFolder.Folders.Add(subFolder.Name);
                destCtx.Load(destSubFolder);
                destCtx.ExecuteQuery();

                _auditLog.Record(AuditAction.FolderCreated, sourceListName, destListName,
                    srcName: subFolder.ServerRelativeUrl,
                    destName: destSubFolder.ServerRelativeUrl);
            }
            else
            {
                _auditLog.Record(AuditAction.FolderAlreadyExists, sourceListName, destListName,
                    srcName: subFolder.ServerRelativeUrl,
                    destName: subFolder.Name,
                    status: "Skipped");
            }

            var freshDstSub = destCtx.Web.GetFolderByServerRelativeUrl(destSubFolder.ServerRelativeUrl);
            destCtx.Load(freshDstSub, f => f.ServerRelativeUrl, f => f.Files, f => f.Folders);
            destCtx.ExecuteQuery();

            await MigrateFilesInFolder(
                sourceCtx, freshSrcSub, destCtx, freshDstSub,
                destList, subRelativePath, destFields,
                sourceListName, destListName, result);
        }
    }
}
