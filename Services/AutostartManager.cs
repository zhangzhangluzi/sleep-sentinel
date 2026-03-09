using Microsoft.Win32;

namespace SleepSentinel.Services;

public static class AutostartManager
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "SleepSentinel";

    public static bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
        var configuredCommand = key?.GetValue(ValueName) as string;
        return string.Equals(
            NormalizeExecutablePath(configuredCommand),
            NormalizeExecutablePath(BuildCommand(Application.ExecutablePath)),
            StringComparison.OrdinalIgnoreCase);
    }

    public static void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);

        if (enabled)
        {
            key?.SetValue(ValueName, BuildCommand(Application.ExecutablePath));
            return;
        }

        key?.DeleteValue(ValueName, false);
    }

    private static string BuildCommand(string executablePath)
    {
        return $"\"{executablePath}\"";
    }

    private static string NormalizeExecutablePath(string? command)
    {
        if (string.IsNullOrWhiteSpace(command))
        {
            return string.Empty;
        }

        var trimmed = command.Trim();
        string executablePath;

        if (trimmed.StartsWith('"'))
        {
            var closingQuoteIndex = trimmed.IndexOf('"', 1);
            executablePath = closingQuoteIndex > 1 ? trimmed[1..closingQuoteIndex] : trimmed.Trim('"');
        }
        else
        {
            var firstSpaceIndex = trimmed.IndexOf(' ');
            executablePath = firstSpaceIndex > 0 ? trimmed[..firstSpaceIndex] : trimmed;
        }

        executablePath = Environment.ExpandEnvironmentVariables(executablePath);

        try
        {
            return Path.GetFullPath(executablePath);
        }
        catch (Exception)
        {
            return executablePath;
        }
    }
}
