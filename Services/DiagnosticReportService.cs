using System.Text;

namespace SleepSentinel.Services;

public sealed class DiagnosticReportService
{
    private readonly SettingsStore _settingsStore;
    private readonly FileLogger _logger;
    private readonly PowerController _controller;

    public DiagnosticReportService(SettingsStore settingsStore, FileLogger logger, PowerController controller)
    {
        _settingsStore = settingsStore;
        _logger = logger;
        _controller = controller;
    }

    public string Export()
    {
        var reportsDirectory = Path.Combine(_settingsStore.BaseDirectory, "reports");
        Directory.CreateDirectory(reportsDirectory);

        var timestamp = DateTime.Now;
        var reportPath = Path.Combine(reportsDirectory, $"diagnostic-report-{timestamp:yyyyMMdd-HHmmss}.txt");
        var settings = _settingsStore.Load();
        var recentLogs = _logger.ReadRecent(400);
        var wakeDiagnostics = _controller.CollectWakeDiagnostics();

        var builder = new StringBuilder();
        builder.AppendLine("SleepSentinel Diagnostic Report");
        builder.AppendLine($"GeneratedAt: {timestamp:yyyy-MM-dd HH:mm:ss}");
        builder.AppendLine();
        builder.AppendLine("[Settings]");
        builder.AppendLine($"PolicyMode: {settings.PolicyMode}");
        builder.AppendLine($"ResumeProtectionEnabled: {settings.ResumeProtectionEnabled}");
        builder.AppendLine($"ResumeProtectionMode: {settings.ResumeProtectionMode}");
        builder.AppendLine($"ResumeProtectionOnlyForUnattendedWake: {settings.ResumeProtectionOnlyForUnattendedWake}");
        builder.AppendLine($"ResumeProtectionDelaySeconds: {settings.ResumeProtectionDelaySeconds}");
        builder.AppendLine($"DisableWakeTimers: {settings.DisableWakeTimers}");
        builder.AppendLine($"WakeTimerPolicySummary: {settings.WakeTimerPolicySummary}");
        builder.AppendLine($"StartMinimized: {settings.StartMinimized}");
        builder.AppendLine($"StartWithWindows: {settings.StartWithWindows}");
        builder.AppendLine($"LastSuspendUtc: {settings.LastSuspendUtc}");
        builder.AppendLine($"LastResumeUtc: {settings.LastResumeUtc}");
        builder.AppendLine($"LastWakeSummary: {settings.LastWakeSummary}");
        builder.AppendLine();
        builder.AppendLine("[Paths]");
        builder.AppendLine($"SettingsPath: {_settingsStore.SettingsPath}");
        builder.AppendLine($"LogDirectory: {_logger.LogDirectory}");
        builder.AppendLine($"ReportPath: {reportPath}");
        builder.AppendLine();
        builder.AppendLine("[WakeDiagnostics]");
        builder.AppendLine(wakeDiagnostics);
        builder.AppendLine();
        builder.AppendLine("[RecentLogs]");

        foreach (var line in recentLogs)
        {
            builder.AppendLine(line);
        }

        File.WriteAllText(reportPath, builder.ToString(), Encoding.UTF8);
        _logger.Info($"已导出诊断报告：{reportPath}");
        return reportPath;
    }
}
