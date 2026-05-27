using System.Text;

namespace MigrationTool2.Models;

public enum AuditAction
{
    ListCreated,
    ListAlreadyExists,
    ListProvisionFailed,
    FieldProvisioned,
    FieldSkipped,
    FieldProvisionFailed,
    ContentTypeProvisioned,
    ContentTypeSkipped,
    FolderCreated,
    FolderAlreadyExists,
    FileUploaded,
    ItemCreated,
    ItemUpdated,
    ItemSkipped,
    ItemFailed,
    AttachmentMigrated,
    DependencyResolved,
    LookupFieldDeferred,
    LookupFieldCreated,
    TaxonomyValueSkipped,
    UserResolved,
    NameVerification
}

public class AuditEntry
{
    public DateTime Timestamp { get; set; }
    public AuditAction Action { get; set; }
    public string SourceListName { get; set; } = string.Empty;
    public string DestinationListName { get; set; } = string.Empty;
    public string? SourceIdentifier { get; set; }
    public string? DestinationIdentifier { get; set; }
    public string? SourceName { get; set; }
    public string? DestinationName { get; set; }
    public string Status { get; set; } = "Success";
    public string? Detail { get; set; }

    public string ToCsv()
    {
        var cols = new[]
        {
            Timestamp.ToString("O"),
            Action.ToString(),
            EscapeCsv(SourceListName),
            EscapeCsv(DestinationListName),
            EscapeCsv(SourceIdentifier ?? ""),
            EscapeCsv(DestinationIdentifier ?? ""),
            EscapeCsv(SourceName ?? ""),
            EscapeCsv(DestinationName ?? ""),
            EscapeCsv(Status),
            EscapeCsv(Detail ?? "")
        };
        return string.Join(",", cols);
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }
}

public class MigrationAuditLog
{
    private readonly string _filePath;
    private readonly object _lock = new();

    private static readonly string Header = "Timestamp,Action,SourceListName,DestinationListName,SourceIdentifier,DestinationIdentifier,SourceName,DestinationName,Status,Detail";

    public MigrationAuditLog(string filePath)
    {
        _filePath = filePath;
        lock (_lock)
        {
            if (!File.Exists(_filePath))
            {
                File.WriteAllLines(_filePath, new[] { Header });
            }
        }
    }

    public void Record(AuditEntry entry)
    {
        lock (_lock)
        {
            var line = entry.ToCsv();
            File.AppendAllLines(_filePath, new[] { line });
        }
    }

    public void Record(AuditAction action, string sourceListName, string destListName,
        string? srcId = null, string? destId = null,
        string? srcName = null, string? destName = null,
        string status = "Success", string? detail = null)
    {
        Record(new AuditEntry
        {
            Timestamp = DateTime.UtcNow,
            Action = action,
            SourceListName = sourceListName,
            DestinationListName = destListName,
            SourceIdentifier = srcId,
            DestinationIdentifier = destId,
            SourceName = srcName,
            DestinationName = destName,
            Status = status,
            Detail = detail
        });
    }

    public void RecordNameVerification(string sourceListName, string destListName,
        string entityType, string sourceName, string destName,
        bool match, string? detail = null)
    {
        Record(AuditAction.NameVerification, sourceListName, destListName,
            srcName: sourceName, destName: destName,
            status: match ? "Match" : "Mismatch",
            detail: $"{entityType}: source='{sourceName}' dest='{destName}'{(detail != null ? " | " + detail : "")}");
    }
}
