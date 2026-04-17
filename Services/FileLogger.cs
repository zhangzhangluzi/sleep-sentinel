using System.Diagnostics;

namespace SleepSentinel.Services;

public sealed class FileLogger
{
    private readonly string _logDirectory;
    private readonly object _sync = new();

    public FileLogger(string baseDirectory)
    {
        _logDirectory = Path.Combine(baseDirectory, "logs");
        Directory.CreateDirectory(_logDirectory);
    }

    public string LogDirectory => _logDirectory;

    public event EventHandler<string>? LogWritten;

    public void Info(string message) => Write("INFO", message);

    public void Warn(string message) => Write("WARN", message);

    public void Error(string message) => Write("ERROR", message);

    public IReadOnlyList<string> ReadRecent(int maxLines = 200)
    {
        if (maxLines <= 0)
        {
            return Array.Empty<string>();
        }

        try
        {
            Directory.CreateDirectory(_logDirectory);

            var remainingLines = maxLines;
            var collectedChunks = new List<string[]>();

            foreach (var path in Directory.EnumerateFiles(_logDirectory, "*.log")
                         .OrderByDescending(Path.GetFileName)
                         .Take(7))
            {
                if (remainingLines <= 0)
                {
                    break;
                }

                var lines = File.ReadAllLines(path).TakeLast(remainingLines).ToArray();
                if (lines.Length == 0)
                {
                    continue;
                }

                collectedChunks.Add(lines);
                remainingLines -= lines.Length;
            }

            if (collectedChunks.Count == 0)
            {
                return Array.Empty<string>();
            }

            collectedChunks.Reverse();
            return collectedChunks.SelectMany(static chunk => chunk).ToArray();
        }
        catch (IOException)
        {
            return Array.Empty<string>();
        }
        catch (UnauthorizedAccessException)
        {
            return Array.Empty<string>();
        }
    }

    private void Write(string level, string message)
    {
        var timestamp = DateTime.Now;
        var line = $"[{timestamp:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
        var path = GetLogPath(timestamp);

        try
        {
            lock (_sync)
            {
                Directory.CreateDirectory(_logDirectory);
                File.AppendAllText(path, line + Environment.NewLine);
            }
        }
        catch (IOException ex)
        {
            Debug.WriteLine($"SleepSentinel log write failed: {ex.Message}");
        }
        catch (UnauthorizedAccessException ex)
        {
            Debug.WriteLine($"SleepSentinel log write denied: {ex.Message}");
        }

        try
        {
            LogWritten?.Invoke(this, line);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"SleepSentinel log subscriber failed: {ex.Message}");
        }
    }

    private string GetLogPath(DateTime timestamp)
    {
        return Path.Combine(_logDirectory, $"{timestamp:yyyy-MM-dd}.log");
    }
}
