using Microsoft.Win32;

namespace SleepSentinel.Services;

public static class AutostartManager
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "SleepSentinel";

    public static bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
        return key?.GetValue(ValueName) is string value && !string.IsNullOrWhiteSpace(value);
    }

    public static void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);

        if (enabled)
        {
            var exePath = Application.ExecutablePath;
            key?.SetValue(ValueName, $"\"{exePath}\"");
            return;
        }

        key?.DeleteValue(ValueName, false);
    }
}
