namespace MigrationTool2.Models;

public class MigrationResult
{
    public int TotalItems { get; set; }
    public int Succeeded { get; set; }
    public int Failed { get; set; }
    public int Skipped { get; set; }
    public List<MigrationError> Errors { get; set; } = new();
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
    public TimeSpan Duration => EndTime - StartTime;
    public string? CurrentList { get; set; }
    public List<string> ListErrors { get; set; } = new();
    public Dictionary<string, ListStats> ListStats { get; set; } = new();

    public void PrintSummary()
    {
        Console.WriteLine($"\n{'='}{'='}{'='} Migration Summary {'='}{'='}{'='}");
        Console.WriteLine($"  Total Items:  {TotalItems}");
        Console.WriteLine($"  Succeeded:    {Succeeded}");
        Console.WriteLine($"  Failed:       {Failed}");
        Console.WriteLine($"  Skipped:      {Skipped}");
        Console.WriteLine($"  Duration:     {Duration:hh\\:mm\\:ss}");
        Console.WriteLine($"  Started:      {StartTime:yyyy-MM-dd HH:mm:ss}");
        Console.WriteLine($"  Completed:    {EndTime:yyyy-MM-dd HH:mm:ss}");

        if (ListStats.Count > 0)
        {
            Console.WriteLine($"\n  Per-List Breakdown:");
            foreach (var (list, stats) in ListStats)
            {
                Console.WriteLine($"    {list}: {stats.Succeeded} ok, {stats.Failed} failed, {stats.Skipped} skipped");
            }
        }

        if (Errors.Count > 0)
        {
            Console.WriteLine($"\n  Errors ({Errors.Count}):");
            foreach (var error in Errors.Take(30))
            {
                Console.WriteLine($"    - [{error.ListName}/{error.ItemId}] {error.Title}: {error.Message}");
            }
            if (Errors.Count > 30)
                Console.WriteLine($"    ... and {Errors.Count - 30} more errors");
        }
        Console.WriteLine(new string('=', 40));
    }
}

public class ListStats
{
    public int Succeeded { get; set; }
    public int Failed { get; set; }
    public int Skipped { get; set; }
}

public class MigrationError
{
    public string ListName { get; set; } = string.Empty;
    public string ItemId { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public Exception? Exception { get; set; }
}
