namespace MigrationTool2.Models;

public class MigrationConfig
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string DestinationSiteUrl { get; set; } = string.Empty;
    public string? SourceListName { get; set; }
    public string? DestinationListName { get; set; }
    public List<string> Lists { get; set; } = new();
    public bool AutoResolveDependencies { get; set; } = true;
    public AuthConfig Auth { get; set; } = new();
    public OptionsConfig Options { get; set; } = new();
}

public class AuthConfig
{
    public string Mode { get; set; } = "DeviceCode";
    public string ClientId { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string CertificateThumbprint { get; set; } = string.Empty;
    public string CertificateFilePath { get; set; } = string.Empty;
    public string CertificatePassword { get; set; } = string.Empty;
}

public class OptionsConfig
{
    public string OverwriteMode { get; set; } = "Skip";
    public bool PreserveCreatedBy { get; set; } = true;
    public bool PreserveModifiedBy { get; set; } = true;
    public bool PreserveCreated { get; set; } = true;
    public bool PreserveModified { get; set; } = true;
    public bool PreserveAttachments { get; set; } = true;
    public bool ResolveLookupDependencies { get; set; } = true;
    public bool ResolveUserDependencies { get; set; } = true;
    public bool MigrateContentTypes { get; set; } = true;
    public int BatchSize { get; set; } = 100;
    public int MaxRetryCount { get; set; } = 3;
}
