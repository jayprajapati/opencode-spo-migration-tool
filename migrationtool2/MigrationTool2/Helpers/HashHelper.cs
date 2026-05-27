using System.Security.Cryptography;

namespace MigrationTool2.Helpers;

public static class HashHelper
{
    public static string ComputeSha256(Stream stream)
    {
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }

    public static string ComputeSha256(byte[] data)
    {
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(data);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }
}
