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
        var path = GetLogPath(DateTime.Now);
        if (!File.Exists(path))
        {
            return Array.Empty<string>();
        }

        return File.ReadAllLines(path).TakeLast(maxLines).ToArray();
    }

    private void Write(string level, string message)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
        var path = GetLogPath(DateTime.Now);

        lock (_sync)
        {
            Directory.CreateDirectory(_logDirectory);
            File.AppendAllText(path, line + Environment.NewLine);
        }

        LogWritten?.Invoke(this, line);
    }

    private string GetLogPath(DateTime timestamp)
    {
        return Path.Combine(_logDirectory, $"{timestamp:yyyy-MM-dd}.log");
    }
}
