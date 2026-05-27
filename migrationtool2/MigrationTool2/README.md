# SharePoint List/Item Migration Tool

Migrates SharePoint list items between sites, preserving regular fields (Title, custom columns) and system metadata (Author, Editor, Created, Modified).

## Requirements

- .NET 7 SDK
- A SharePoint Online tenant with two site collections (source + destination)
- Azure AD app registration with appropriate permissions (see [Auth Setup](#auth-setup))

## Quick Start

```bash
# 1. Configure appsettings.json with your sites and list name
# 2. Run a full migration + verify
dotnet run -- --clean       # delete destination list (fresh start)
dotnet run                  # migrate items
dotnet run -- --verify      # compare source vs destination
```

## CLI Flags

| Flag | Description |
|------|-------------|
| *(no flag)* | Run migration normally |
| `--clean` | Delete destination list before migration |
| `--verify` | Compare source/dest items (Created, Modified, Author) |
| `--test-rest` | Ad-hoc REST API diagnostic tests |

## Configuration (`appsettings.json`)

```jsonc
{
  "SourceSiteUrl": "https://tenant.sharepoint.com/sites/source",
  "DestinationSiteUrl": "https://tenant.sharepoint.com/sites/dest",
  "SourceListName": "Cars",
  "DestinationListName": "Cars",
  "Lists": [],                       // Optional: explicit list of lists to migrate
  "AutoResolveDependencies": true,   // Resolve lookup column dependencies between lists
  "Auth": { /* see Auth Setup below */ },
  "Options": {
    "OverwriteMode": "Overwrite",    // "Overwrite" | "Skip"
    "PreserveCreatedBy": true,
    "PreserveModifiedBy": true,
    "PreserveCreated": true,
    "PreserveModified": true,
    "PreserveAttachments": true,
    "ResolveLookupDependencies": true,
    "ResolveUserDependencies": true,
    "MigrateContentTypes": true,
    "BatchSize": 100,
    "MaxRetryCount": 3
  }
}
```

## Auth Setup

### Certificate (recommended for automation)

1. Register an app in [Azure AD](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/):
   - API permissions: `Sites.ReadWrite.All` (application type)
   - Upload a self-signed certificate (public key)
2. Configure:

```jsonc
"Auth": {
  "Mode": "Certificate",                // or "AppOnly"
  "ClientId": "11111111-...",
  "TenantId": "22222222-...",
  "CertificateFilePath": ".certs/dev-cert.pfx",
  "CertificatePassword": ""             // empty if no password
}
```

3. Place the `.pfx` file in `.certs/`. The build copies it to the output directory automatically.

### Client Secret (alternative server-to-server)

```jsonc
"Auth": {
  "Mode": "AppOnly",
  "ClientId": "11111111-...",
  "TenantId": "22222222-...",
  "ClientSecret": "your-secret"
}
```

### Device Code (interactive, user-delegated)

```jsonc
"Auth": {
  "Mode": "DeviceCode",
  "ClientId": "11111111-...",
  "TenantId": "22222222-..."
}
```

Opens a browser prompt for user login at runtime.

## How It Works

1. **Schema Migration** — Creates destination list + provisions all fields from source
2. **Item Migration** — Reads source items via CSOM, writes to destination via SharePoint REST API (bypasses CSOM app-only field-write limitation)
3. **System Field Preservation** — Uses `ValidateUpdateListItem` with `bNewDocumentUpdate=true` to set Author, Editor, Created, Modified dates
4. **Timezone Handling** — Queries the SharePoint site's timezone settings and converts UTC source dates to site-local time before writing
5. **Verification** — Reads both source and destination items and compares Created/Modified dates (with tolerance for seconds truncation) and field values

## What's Migrated

- All non-hidden, non-readonly, non-system fields (Text, Number, DateTime, Choice, Lookup, User, URL, Taxonomy, Boolean, etc.)
- Title field
- System metadata (Author, Editor, Created, Modified) — when `Preserve*` options are enabled
- Content types
- Attachments

## Known Limitations

- **Seconds truncation**: SharePoint's `ValidateUpdateListItem` truncates Created/Modified seconds to `00`. The verify tolerates this with a 70-second margin.
- **AuthorId differs across sites**: The same user has different integer IDs on different site collections. The verify notes this but doesn't fail.
- **Lookup dependencies**: Lists with lookup columns referencing other lists need `AutoResolveDependencies: true` and the referenced list should either already exist on the destination or be included in the migration order.
