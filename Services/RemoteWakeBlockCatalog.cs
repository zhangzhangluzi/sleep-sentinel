using System.Diagnostics;
using System.Text.RegularExpressions;

namespace SleepSentinel.Services;

internal enum PowerRequestOverrideCallerType
{
    Process,
    Service
}

internal readonly record struct PowerRequestOverrideEntry(
    PowerRequestOverrideCallerType CallerType,
    string Name,
    string Product)
{
    public string CallerTypeArgument => CallerType == PowerRequestOverrideCallerType.Process ? "process" : "service";
}

internal static class RemoteWakeBlockCatalog
{
    private static readonly Regex RequestEntryRegex = new(@"^\[(PROCESS|SERVICE)\]\s*(?<value>.+)$", RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly string[] DiscoveryKeywords =
    [
        "anydesk",
        "awesun",
        "gameviewer",
        "parsec",
        "remote",
        "rustdesk",
        "splashtop",
        "sunlogin",
        "teamviewer",
        "todesk",
        "ultraviewer",
        "uu",
        "vnc"
    ];

    public static IReadOnlyList<PowerRequestOverrideEntry> BuiltInEntries { get; } =
    [
        new(PowerRequestOverrideCallerType.Process, "ToDesk.exe", "ToDesk"),
        new(PowerRequestOverrideCallerType.Process, "CoreToolsMgrHelper.exe", "ToDesk"),
        new(PowerRequestOverrideCallerType.Service, "ToDesk_Service", "ToDesk"),

        new(PowerRequestOverrideCallerType.Process, "SunloginClient.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "SunloginRemote.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "SunloginService.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "AweSun.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "OldAweSun.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Service, "SunloginService", "向日葵 / AweSun"),

        new(PowerRequestOverrideCallerType.Process, "GameViewer.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerLauncher.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerServer.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerHealthd.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerService.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Service, "GameViewerService", "GameViewer / UU远控"),

        new(PowerRequestOverrideCallerType.Process, "AnyDesk.exe", "AnyDesk"),
        new(PowerRequestOverrideCallerType.Process, "ad_svc.exe", "AnyDesk"),
        new(PowerRequestOverrideCallerType.Service, "AnyDesk", "AnyDesk"),

        new(PowerRequestOverrideCallerType.Process, "TeamViewer.exe", "TeamViewer"),
        new(PowerRequestOverrideCallerType.Process, "TeamViewer_Service.exe", "TeamViewer"),
        new(PowerRequestOverrideCallerType.Service, "TeamViewer", "TeamViewer"),

        new(PowerRequestOverrideCallerType.Process, "RustDesk.exe", "RustDesk"),
        new(PowerRequestOverrideCallerType.Process, "rustdesk.exe", "RustDesk"),
        new(PowerRequestOverrideCallerType.Process, "RustDeskService.exe", "RustDesk"),
        new(PowerRequestOverrideCallerType.Service, "RustDesk", "RustDesk")
    ];

    public static IReadOnlyList<PowerRequestOverrideEntry> GetManagedEntries(IEnumerable<string>? customRawEntries)
    {
        var entries = new Dictionary<string, PowerRequestOverrideEntry>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in BuiltInEntries)
        {
            entries[BuildDedupKey(entry.CallerType, entry.Name)] = entry;
        }

        foreach (var entry in ParseCustomEntries(customRawEntries))
        {
            entries[BuildDedupKey(entry.CallerType, entry.Name)] = entry;
        }

        return entries.Values.ToArray();
    }

    public static IReadOnlyList<string> NormalizeCustomEntries(IEnumerable<string>? rawEntries)
    {
        return ParseCustomEntries(rawEntries)
            .Select(FormatCustomEntry)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static entry => entry, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public static IReadOnlyList<string> SuggestCustomEntries(string powerRequestsText, IEnumerable<string>? existingCustomEntries = null)
    {
        var suggestions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var existingManagedKeys = new HashSet<string>(
            GetManagedEntries(existingCustomEntries).Select(static entry => BuildDedupKey(entry.CallerType, entry.Name)),
            StringComparer.OrdinalIgnoreCase);

        foreach (Match match in RequestEntryRegex.Matches(powerRequestsText))
        {
            var type = match.Groups[1].Value.Equals("SERVICE", StringComparison.OrdinalIgnoreCase)
                ? PowerRequestOverrideCallerType.Service
                : PowerRequestOverrideCallerType.Process;
            var value = match.Groups["value"].Value.Trim();
            var normalizedName = type == PowerRequestOverrideCallerType.Process
                ? ExtractProcessName(value)
                : value;

            if (TryCreateSuggestion(type, normalizedName, out var suggestion)
                && existingManagedKeys.Add(BuildDedupKey(suggestion.CallerType, suggestion.Name)))
            {
                suggestions.Add(FormatCustomEntry(suggestion));
            }
        }

        foreach (var process in Process.GetProcesses())
        {
            try
            {
                var processName = process.ProcessName;
                if (string.IsNullOrWhiteSpace(processName))
                {
                    continue;
                }

                var normalizedName = processName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                    ? processName
                    : $"{processName}.exe";

                if (TryCreateSuggestion(PowerRequestOverrideCallerType.Process, normalizedName, out var suggestion)
                    && existingManagedKeys.Add(BuildDedupKey(suggestion.CallerType, suggestion.Name)))
                {
                    suggestions.Add(FormatCustomEntry(suggestion));
                }
            }
            catch (InvalidOperationException)
            {
            }
        }

        return suggestions.OrderBy(static entry => entry, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public static string ProductSummary(IEnumerable<string>? customRawEntries = null)
    {
        var products = BuiltInEntries
            .Select(static entry => entry.Product)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var customCount = ParseCustomEntries(customRawEntries).Count;
        if (customCount > 0)
        {
            products.Add($"自定义 {customCount} 条");
        }

        return string.Join("、", products);
    }

    private static List<PowerRequestOverrideEntry> ParseCustomEntries(IEnumerable<string>? rawEntries)
    {
        var entries = new List<PowerRequestOverrideEntry>();
        if (rawEntries is null)
        {
            return entries;
        }

        foreach (var rawLine in rawEntries)
        {
            var line = rawLine?.Trim();
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (!TryParseCustomEntry(line, out var entry))
            {
                continue;
            }

            entries.Add(entry);
        }

        return entries;
    }

    private static bool TryParseCustomEntry(string line, out PowerRequestOverrideEntry entry)
    {
        entry = default;
        var normalized = line.Trim().Trim('"');

        string callerTypeToken;
        string name;

        var separatorIndex = normalized.IndexOf(':');
        if (separatorIndex > 0)
        {
            callerTypeToken = normalized[..separatorIndex].Trim();
            name = normalized[(separatorIndex + 1)..].Trim().Trim('"');
        }
        else
        {
            callerTypeToken = normalized.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) ? "process" : "service";
            name = normalized;
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        if (callerTypeToken.Equals("process", StringComparison.OrdinalIgnoreCase))
        {
            if (!name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                name += ".exe";
            }

            entry = new PowerRequestOverrideEntry(PowerRequestOverrideCallerType.Process, name, "自定义");
            return true;
        }

        if (callerTypeToken.Equals("service", StringComparison.OrdinalIgnoreCase))
        {
            entry = new PowerRequestOverrideEntry(PowerRequestOverrideCallerType.Service, name, "自定义");
            return true;
        }

        return false;
    }

    private static bool TryCreateSuggestion(PowerRequestOverrideCallerType callerType, string name, out PowerRequestOverrideEntry entry)
    {
        entry = default;
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        var normalized = name.Trim();
        if (!ContainsDiscoveryKeyword(normalized))
        {
            return false;
        }

        if (callerType == PowerRequestOverrideCallerType.Process && !normalized.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            normalized += ".exe";
        }

        entry = new PowerRequestOverrideEntry(callerType, normalized, "自定义");
        return true;
    }

    private static bool ContainsDiscoveryKeyword(string value)
    {
        return DiscoveryKeywords.Any(keyword => value.Contains(keyword, StringComparison.OrdinalIgnoreCase));
    }

    private static string ExtractProcessName(string value)
    {
        var trimmed = value.Trim().Trim('"');
        if (trimmed.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            return Path.GetFileName(trimmed);
        }

        return Path.GetFileName(trimmed) ?? trimmed;
    }

    private static string FormatCustomEntry(PowerRequestOverrideEntry entry)
    {
        return $"{entry.CallerTypeArgument}:{entry.Name}";
    }

    private static string BuildDedupKey(PowerRequestOverrideCallerType callerType, string name)
    {
        return $"{callerType}:{name}";
    }
}
