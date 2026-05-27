using Microsoft.SharePoint.Client;
using MigrationTool2.Helpers;
using MigrationTool2.Models;

namespace MigrationTool2.Services;

public class AttachmentMigrationService
{
    private readonly ThrottleHandler _throttle;
    private readonly MigrationAuditLog _auditLog;

    public AttachmentMigrationService(ThrottleHandler throttle, MigrationAuditLog auditLog)
    {
        _throttle = throttle;
        _auditLog = auditLog;
    }

    public async Task MigrateAttachments(
        ClientContext sourceCtx, ListItem sourceItem,
        ClientContext destCtx, ListItem destItem,
        string sourceListName, MigrationResult result,
        string? destListRootUrl = null)
    {
        sourceCtx.Load(sourceItem, i => i.AttachmentFiles);
        sourceCtx.ExecuteQuery();

        if (sourceItem.AttachmentFiles.Count == 0) return;

        foreach (var attachment in sourceItem.AttachmentFiles)
        {
            sourceCtx.Load(attachment, a => a.FileName, a => a.ServerRelativeUrl);
            sourceCtx.ExecuteQuery();

            try
            {
                await _throttle.ExecuteWithRetryAsync(async () =>
                {
                    var attFile = sourceCtx.Web.GetFileByServerRelativeUrl(attachment.ServerRelativeUrl);
                    sourceCtx.Load(attFile, f => f.Name);
                    var streamRefTask = attFile.OpenBinaryStream();
                    sourceCtx.ExecuteQuery();

                    var attachmentInfo = new AttachmentCreationInformation
                    {
                        FileName = attachment.FileName,
                        ContentStream = streamRefTask.Value
                    };

                    try
                    {
                        var newAttachment = destItem.AttachmentFiles.Add(attachmentInfo);
                        destCtx.Load(newAttachment);
                        destCtx.ExecuteQuery();
                    }
                    catch (Exception addEx) when (addEx.Message.Contains("already exists"))
                    {
                        await DeleteExistingAttachment(destCtx, destItem, attachment.FileName, destListRootUrl);
                        var retryAttachment = destItem.AttachmentFiles.Add(attachmentInfo);
                        destCtx.Load(retryAttachment);
                        destCtx.ExecuteQuery();
                    }

                    _auditLog.Record(AuditAction.AttachmentMigrated, sourceListName, sourceListName,
                        srcId: sourceItem.Id.ToString(), destId: destItem.Id.ToString(),
                        srcName: attachment.FileName, destName: attachment.FileName);
                }, $"Attachment: {attachment.FileName}");
            }
            catch (Exception ex)
            {
                _auditLog.Record(AuditAction.AttachmentMigrated, sourceListName, sourceListName,
                    srcId: sourceItem.Id.ToString(), destId: destItem.Id.ToString(),
                    srcName: attachment.FileName,
                    status: "Failed", detail: ex.Message);

                result.Errors.Add(new MigrationError
                {
                    ListName = sourceListName,
                    ItemId = $"{sourceItem.Id}:{attachment.FileName}",
                    Title = attachment.FileName,
                    Message = $"Attachment migration failed: {ex.Message}",
                    Exception = ex
                });
                Console.WriteLine($"  ! Attachment FAILED: {attachment.FileName} - {ex.Message}");
            }
        }
    }

    private async Task DeleteExistingAttachment(ClientContext ctx, ListItem item, string fileName, string? listRootUrl)
    {
        if (listRootUrl == null) return;

        var fileUrl = listRootUrl.TrimEnd('/') + $"/Attachments/{item.Id}/{fileName}";

        try
        {
            var file = ctx.Web.GetFileByServerRelativeUrl(fileUrl);
            ctx.Load(file);
            await _throttle.ExecuteWithRetryAsync(() =>
            {
                ctx.ExecuteQuery();
                return Task.CompletedTask;
            }, $"Delete existing attachment: {fileName}");

            file.DeleteObject();
            await _throttle.ExecuteWithRetryAsync(() =>
            {
                ctx.ExecuteQuery();
                return Task.CompletedTask;
            }, $"Delete existing attachment (execute): {fileName}");
        }
        catch
        {
        }
    }
}
