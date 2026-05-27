using System.Net;
using Microsoft.SharePoint.Client;

namespace MigrationTool2.Helpers;

public class ThrottleHandler
{
    private readonly int _maxRetries;

    public ThrottleHandler(int maxRetries = 3)
    {
        _maxRetries = maxRetries;
    }

    public async Task ExecuteWithRetryAsync(Func<Task> action, string context = "")
    {
        int retryCount = 0;
        while (true)
        {
            try
            {
                await action();
                return;
            }
            catch (ServerUnauthorizedAccessException)
            {
                throw;
            }
            catch (WebException ex) when (ex.Response is HttpWebResponse response && response.StatusCode == (HttpStatusCode)429)
            {
                retryCount++;
                if (retryCount > _maxRetries)
                    throw new InvalidOperationException($"Throttled too many times: {context}", ex);

                var retryAfter = response.Headers["Retry-After"];
                var delayMs = retryAfter != null && int.TryParse(retryAfter, out var seconds)
                    ? seconds * 1000
                    : (int)Math.Pow(2, retryCount) * 1000;

                Console.WriteLine($"  Throttled. Retry {retryCount}/{_maxRetries} after {delayMs}ms. Context: {context}");
                await Task.Delay(delayMs);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                retryCount++;
                if (retryCount > _maxRetries)
                    throw;

                var delay = (int)Math.Pow(2, retryCount) * 1000;
                Console.WriteLine($"  Transient error: {ex.Message}. Retry {retryCount}/{_maxRetries} after {delay}ms.");
                await Task.Delay(delay);
            }
        }
    }

    public void ExecuteWithRetry(Action action, string context = "")
    {
        ExecuteWithRetryAsync(() => { action(); return Task.CompletedTask; }, context).GetAwaiter().GetResult();
    }
}
