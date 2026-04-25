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
        var reportPath = Path.Combine(reportsDirectory, $"diagnostic-report-{timestamp:yyyyMMdd-HHmmss-fff}-{Guid.NewGuid():N}.txt");
        var tempPath = reportPath + ".tmp";
        var settings = _controller.CurrentSettings;
        var recentLogs = _logger.ReadRecent(400);
        var snapshot = _controller.CollectWakeDiagnosticSnapshot(includePowerRequests: true, includeSleepStudy: true);
        var wakeDiagnostics = _controller.FormatWakeDiagnosticSnapshot(snapshot, includePowerRequests: false, includeSleepStudy: false);
        var powerRequestDiagnostics =
            $"requests:{Environment.NewLine}{snapshot.PowerRequestsText}{Environment.NewLine}{Environment.NewLine}requestsoverride:{Environment.NewLine}{snapshot.RequestOverridesText}";

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
        builder.AppendLine($"DisableWiFiDirectAdapters: {settings.DisableWiFiDirectAdapters}");
        builder.AppendLine($"EnforceBatteryStandbyHibernate: {settings.EnforceBatteryStandbyHibernate}");
        builder.AppendLine($"BatteryStandbyHibernateTimeoutSeconds: {settings.BatteryStandbyHibernateTimeoutSeconds}");
        builder.AppendLine($"BlockKnownRemoteWakeRequests: {settings.BlockKnownRemoteWakeRequests}");
        builder.AppendLine($"MonitorRayLinkProcessStorm: {settings.MonitorRayLinkProcessStorm}");
        builder.AppendLine($"AutoContainRayLinkProcessStorm: {settings.AutoContainRayLinkProcessStorm}");
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
        builder.AppendLine($"WiFiDirectAdapterRestoreSnapshotCaptured: {settings.WiFiDirectAdapterRestoreSnapshotCaptured}");
        builder.AppendLine($"WiFiDirectAdapterRestoreInstanceIds: {string.Join(", ", settings.WiFiDirectAdapterRestoreInstanceIds)}");
        builder.AppendLine($"WiFiDirectAdapterPolicySummary: {settings.WiFiDirectAdapterPolicySummary}");
        builder.AppendLine($"BatteryStandbyHibernateRestoreSnapshotCaptured: {settings.BatteryStandbyHibernateRestoreSnapshotCaptured}");
        builder.AppendLine($"BatteryStandbyHibernateRestoreAcValue: {settings.BatteryStandbyHibernateRestoreAcValue}");
        builder.AppendLine($"BatteryStandbyHibernateRestoreDcValue: {settings.BatteryStandbyHibernateRestoreDcValue}");
        builder.AppendLine($"BatteryStandbyHibernatePolicySummary: {settings.BatteryStandbyHibernatePolicySummary}");
        builder.AppendLine($"KnownRemoteWakeRequestBackupCaptured: {settings.KnownRemoteWakeRequestBackupCaptured}");
        builder.AppendLine($"KnownRemoteWakeRequestOverrideBackupCount: {settings.KnownRemoteWakeRequestOverrideBackup.Count}");
        builder.AppendLine($"KnownRemoteWakePolicySummary: {settings.KnownRemoteWakePolicySummary}");
        builder.AppendLine($"RayLinkProcessStormPolicySummary: {settings.RayLinkProcessStormPolicySummary}");
        builder.AppendLine($"AutostartPolicySummary: {settings.AutostartPolicySummary}");
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

        try
        {
            File.WriteAllText(tempPath, builder.ToString(), Encoding.UTF8);
            File.Move(tempPath, reportPath, true);
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException($"写入诊断报告失败：{ex.Message}", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new InvalidOperationException($"写入诊断报告被拒绝：{ex.Message}", ex);
        }
        finally
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        _logger.Info($"已导出诊断报告：{reportPath}");
        return reportPath;
    }
}
