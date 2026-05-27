using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;

namespace MigrationTool2.Services;

public class AuthService
{
    private readonly Models.AuthConfig _auth;
    private string? _cachedToken;
    private DateTime _tokenExpiry;

    public AuthService(Models.AuthConfig auth)
    {
        _auth = auth;
    }

    public ClientContext GetContext(string siteUrl)
    {
        var ctx = new ClientContext(siteUrl);

        ctx.ExecutingWebRequest += (sender, e) =>
        {
            var token = GetAccessTokenAsync(siteUrl).GetAwaiter().GetResult();
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
        };

        return ctx;
    }

    public async Task<string> GetAccessTokenAsync(string siteUrl)
    {
        if (!string.IsNullOrEmpty(_cachedToken) && DateTime.UtcNow < _tokenExpiry)
            return _cachedToken;

        var rootUrl = ExtractRootUrl(siteUrl);
        var scopes = new[] { $"{rootUrl}/.default" };

        if (_auth.Mode is "AppOnly" or "Certificate")
        {
            return await GetAppOnlyTokenAsync(scopes);
        }

        return await GetDeviceCodeTokenAsync(scopes);
    }

    private async Task<string> GetAppOnlyTokenAsync(string[] scopes)
    {
        var appBuilder = ConfidentialClientApplicationBuilder
            .Create(_auth.ClientId)
            .WithTenantId(_auth.TenantId);

        var certPath = _auth.CertificateFilePath;
        if (!string.IsNullOrEmpty(certPath) && !Path.IsPathRooted(certPath))
        {
            var asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;
            certPath = Path.GetFullPath(Path.Combine(asmDir, certPath));
        }

        if (!string.IsNullOrEmpty(certPath) && System.IO.File.Exists(certPath))
        {
            var certBytes = System.IO.File.ReadAllBytes(certPath);
            System.Security.Cryptography.X509Certificates.X509Certificate2 cert;
            if (string.IsNullOrEmpty(_auth.CertificatePassword))
                cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(certBytes);
            else
            {
                var securePwd = new System.Security.SecureString();
                foreach (var c in _auth.CertificatePassword.ToCharArray())
                    securePwd.AppendChar(c);
                cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(certBytes, securePwd);
            }
            appBuilder = appBuilder.WithCertificate(cert);
        }
        else if (!string.IsNullOrEmpty(_auth.CertificateThumbprint))
        {
            using var store = new System.Security.Cryptography.X509Certificates.X509Store(
                System.Security.Cryptography.X509Certificates.StoreName.My);
            store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly);
            var certs = store.Certificates.Find(
                System.Security.Cryptography.X509Certificates.X509FindType.FindByThumbprint,
                _auth.CertificateThumbprint, false);
            if (certs.Count > 0)
                appBuilder = appBuilder.WithCertificate(certs[0]);
        }
        else if (!string.IsNullOrEmpty(_auth.ClientSecret))
        {
            appBuilder = appBuilder.WithClientSecret(_auth.ClientSecret);
        }

        var app = appBuilder.Build();
        var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
        CacheToken(result);
        return _cachedToken!;
    }

    private async Task<string> GetDeviceCodeTokenAsync(string[] scopes)
    {
        var app = PublicClientApplicationBuilder
            .Create(_auth.ClientId)
            .WithTenantId(_auth.TenantId)
            .WithDefaultRedirectUri()
            .Build();

        var accounts = await app.GetAccountsAsync();
        try
        {
            var result = await app.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                .ExecuteAsync();
            CacheToken(result);
            return _cachedToken!;
        }
        catch (MsalUiRequiredException)
        {
            var result = await app.AcquireTokenWithDeviceCode(scopes, deviceCodeResult =>
            {
                Console.WriteLine("\n====================================");
                Console.WriteLine("  DEVICE CODE AUTHENTICATION REQUIRED");
                Console.WriteLine("====================================");
                Console.WriteLine($"  Open: {deviceCodeResult.VerificationUrl}");
                Console.WriteLine($"  Code:  {deviceCodeResult.UserCode}");
                Console.WriteLine("====================================\n");
                return Task.CompletedTask;
            }).ExecuteAsync();

            CacheToken(result);
            return _cachedToken!;
        }
    }

    private void CacheToken(AuthenticationResult result)
    {
        _cachedToken = result.AccessToken;
        _tokenExpiry = result.ExpiresOn.DateTime;
    }

    private static string ExtractRootUrl(string siteUrl)
    {
        var uri = new Uri(siteUrl);
        return $"{uri.Scheme}://{uri.Host}";
    }
}
