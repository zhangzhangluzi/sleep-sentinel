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
        var snapshot = _controller.CollectWakeDiagnosticSnapshot(includePowerRequests: true, includeSleepStudy: true);
        var wakeDiagnostics = _controller.FormatWakeDiagnosticSnapshot(snapshot, includePowerRequests: false, includeSleepStudy: false);
        var powerRequestDiagnostics = _controller.CollectPowerRequestDiagnostics();

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
        builder.AppendLine($"DisableStandbyConnectivity: {settings.DisableStandbyConnectivity}");
        builder.AppendLine($"EnforceBatteryStandbyHibernate: {settings.EnforceBatteryStandbyHibernate}");
        builder.AppendLine($"BatteryStandbyHibernateTimeoutSeconds: {settings.BatteryStandbyHibernateTimeoutSeconds}");
        builder.AppendLine($"BlockKnownRemoteWakeRequests: {settings.BlockKnownRemoteWakeRequests}");
        builder.AppendLine($"CustomRemoteWakeEntries: {string.Join(", ", settings.CustomRemoteWakeEntries)}");
        builder.AppendLine($"ActivePowerPlanSummary: {_controller.CurrentPowerPlanSummary}");
        builder.AppendLine($"RiskSummary: {_controller.CurrentRiskSummary}");
        builder.AppendLine($"CapabilitySummary: {_controller.CurrentCapabilitySummary}");
        builder.AppendLine($"ManagedRemoteEntriesSummary: {_controller.CurrentManagedRemoteEntriesSummary}");
        builder.AppendLine($"WakeTimerRestoreSnapshotCaptured: {settings.WakeTimerRestoreSnapshotCaptured}");
        builder.AppendLine($"WakeTimerRestoreAcValue: {settings.WakeTimerRestoreAcValue}");
        builder.AppendLine($"WakeTimerRestoreDcValue: {settings.WakeTimerRestoreDcValue}");
        builder.AppendLine($"WakeTimerPolicySummary: {settings.WakeTimerPolicySummary}");
        builder.AppendLine($"StandbyConnectivityRestoreSnapshotCaptured: {settings.StandbyConnectivityRestoreSnapshotCaptured}");
        builder.AppendLine($"StandbyConnectivityRestoreAcValue: {settings.StandbyConnectivityRestoreAcValue}");
        builder.AppendLine($"StandbyConnectivityRestoreDcValue: {settings.StandbyConnectivityRestoreDcValue}");
        builder.AppendLine($"DisconnectedStandbyModeRestoreAcValue: {settings.DisconnectedStandbyModeRestoreAcValue}");
        builder.AppendLine($"DisconnectedStandbyModeRestoreDcValue: {settings.DisconnectedStandbyModeRestoreDcValue}");
        builder.AppendLine($"StandbyConnectivityPolicySummary: {settings.StandbyConnectivityPolicySummary}");
        builder.AppendLine($"BatteryStandbyHibernateRestoreSnapshotCaptured: {settings.BatteryStandbyHibernateRestoreSnapshotCaptured}");
        builder.AppendLine($"BatteryStandbyHibernateRestoreAcValue: {settings.BatteryStandbyHibernateRestoreAcValue}");
        builder.AppendLine($"BatteryStandbyHibernateRestoreDcValue: {settings.BatteryStandbyHibernateRestoreDcValue}");
        builder.AppendLine($"BatteryStandbyHibernatePolicySummary: {settings.BatteryStandbyHibernatePolicySummary}");
        builder.AppendLine($"KnownRemoteWakeRequestBackupCaptured: {settings.KnownRemoteWakeRequestBackupCaptured}");
        builder.AppendLine($"KnownRemoteWakeRequestOverrideBackupCount: {settings.KnownRemoteWakeRequestOverrideBackup.Count}");
        builder.AppendLine($"KnownRemoteWakePolicySummary: {settings.KnownRemoteWakePolicySummary}");
        builder.AppendLine($"ProtectionRuleSummary: {_controller.CurrentProtectionRuleSummary}");
        builder.AppendLine($"StartMinimized: {settings.StartMinimized}");
        builder.AppendLine($"StartWithWindows: {settings.StartWithWindows}");
        builder.AppendLine($"LastSuspendUtc: {settings.LastSuspendUtc}");
        builder.AppendLine($"LastResumeUtc: {settings.LastResumeUtc}");
        builder.AppendLine($"LastWakeSummary: {settings.LastWakeSummary}");
        builder.AppendLine($"LastWakeEvidenceSummary: {settings.LastWakeEvidenceSummary}");
        builder.AppendLine();
        builder.AppendLine("[Paths]");
        builder.AppendLine($"SettingsPath: {_settingsStore.SettingsPath}");
        builder.AppendLine($"LogDirectory: {_logger.LogDirectory}");
        builder.AppendLine($"ReportPath: {reportPath}");
        builder.AppendLine();
        builder.AppendLine("[WakeDiagnostics]");
        builder.AppendLine(wakeDiagnostics);
        builder.AppendLine();
        builder.AppendLine("[PowerRequestDiagnostics]");
        builder.AppendLine(powerRequestDiagnostics);
        builder.AppendLine();
        builder.AppendLine("[SleepStudy]");
        builder.AppendLine(snapshot.SleepStudySummary);
        builder.AppendLine();
        builder.AppendLine("[EventSummary]");
        builder.AppendLine(snapshot.EventSummary);
        builder.AppendLine();
        builder.AppendLine("[RemoteWakeSuggestions]");
        builder.AppendLine(snapshot.SuggestedRemoteWakeEntriesSummary);
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
