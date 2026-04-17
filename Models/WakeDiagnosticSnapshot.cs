namespace SleepSentinel.Models;

public sealed class WakeDiagnosticSnapshot
{
    public DateTimeOffset GeneratedAtUtc { get; set; } = DateTimeOffset.UtcNow;
    public string LastWakeText { get; set; } = string.Empty;
    public string WakeTimersText { get; set; } = string.Empty;
    public string PowerRequestsText { get; set; } = string.Empty;
    public string RequestOverridesText { get; set; } = string.Empty;
    public string EventSummary { get; set; } = string.Empty;
    public string SleepStudySummary { get; set; } = string.Empty;
    public IReadOnlyList<string> SuggestedRemoteWakeEntries { get; set; } = [];
    public string SuggestedRemoteWakeEntriesSummary { get; set; } = string.Empty;
}
