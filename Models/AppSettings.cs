namespace SleepSentinel.Models;

public sealed class AppSettings
{
    public PowerPolicyMode PolicyMode { get; set; } = PowerPolicyMode.FollowPowerPlan;
    public ResumeProtectionMode ResumeProtectionMode { get; set; } = ResumeProtectionMode.Hibernate;
    public bool ResumeProtectionEnabled { get; set; } = true;
    public bool ResumeProtectionOnlyForUnattendedWake { get; set; } = true;
    public bool DisableWakeTimers { get; set; }
    public bool DisableStandbyConnectivity { get; set; }
    public bool DisableWiFiDirectAdapters { get; set; }
    public bool EnforceBatteryStandbyHibernate { get; set; }
    public bool BlockKnownRemoteWakeRequests { get; set; }
    public int BatteryStandbyHibernateTimeoutSeconds { get; set; } = 600;
    public Dictionary<string, PowerPlanRestoreSnapshot> PowerPlanRestoreSnapshots { get; set; } = [];
    public List<string> CustomRemoteWakeEntries { get; set; } = [];
    public bool WakeTimerRestoreSnapshotCaptured { get; set; }
    public int WakeTimerRestoreAcValue { get; set; }
    public int WakeTimerRestoreDcValue { get; set; }
    public bool StandbyConnectivityRestoreSnapshotCaptured { get; set; }
    public int StandbyConnectivityRestoreAcValue { get; set; }
    public int StandbyConnectivityRestoreDcValue { get; set; }
    public int DisconnectedStandbyModeRestoreAcValue { get; set; }
    public int DisconnectedStandbyModeRestoreDcValue { get; set; }
    public bool BatteryStandbyHibernateRestoreSnapshotCaptured { get; set; }
    public int BatteryStandbyHibernateRestoreAcValue { get; set; }
    public int BatteryStandbyHibernateRestoreDcValue { get; set; }
    public bool WiFiDirectAdapterRestoreSnapshotCaptured { get; set; }
    public List<string> WiFiDirectAdapterRestoreInstanceIds { get; set; } = [];
    public bool KnownRemoteWakeRequestBackupCaptured { get; set; }
    public Dictionary<string, int> KnownRemoteWakeRequestOverrideBackup { get; set; } = [];
    public int ResumeProtectionDelaySeconds { get; set; } = 8;
    public bool StartMinimized { get; set; } = true;
    public bool StartWithWindows { get; set; }
    public bool WindowBoundsCaptured { get; set; }
    public int WindowWidth { get; set; }
    public int WindowHeight { get; set; }
    public int WindowX { get; set; }
    public int WindowY { get; set; }
    public DateTimeOffset? LastSuspendUtc { get; set; }
    public DateTimeOffset? LastResumeUtc { get; set; }
    public string LastWakeSummary { get; set; } = "无";
    public string LastWakeEvidenceSummary { get; set; } = "无";
    public string WakeTimerPolicySummary { get; set; } = "未检查";
    public string StandbyConnectivityPolicySummary { get; set; } = "未检查";
    public string WiFiDirectAdapterPolicySummary { get; set; } = "未检查";
    public string BatteryStandbyHibernatePolicySummary { get; set; } = "未检查";
    public string KnownRemoteWakePolicySummary { get; set; } = "未检查";
    public string AutostartPolicySummary { get; set; } = "未检查";

    public AppSettings Clone()
    {
        return new AppSettings
        {
            PolicyMode = PolicyMode,
            ResumeProtectionMode = ResumeProtectionMode,
            ResumeProtectionEnabled = ResumeProtectionEnabled,
            ResumeProtectionOnlyForUnattendedWake = ResumeProtectionOnlyForUnattendedWake,
            DisableWakeTimers = DisableWakeTimers,
            DisableStandbyConnectivity = DisableStandbyConnectivity,
            DisableWiFiDirectAdapters = DisableWiFiDirectAdapters,
            EnforceBatteryStandbyHibernate = EnforceBatteryStandbyHibernate,
            BlockKnownRemoteWakeRequests = BlockKnownRemoteWakeRequests,
            BatteryStandbyHibernateTimeoutSeconds = BatteryStandbyHibernateTimeoutSeconds,
            PowerPlanRestoreSnapshots = PowerPlanRestoreSnapshots.ToDictionary(
                static pair => pair.Key,
                static pair => pair.Value.Clone(),
                StringComparer.OrdinalIgnoreCase),
            CustomRemoteWakeEntries = [.. CustomRemoteWakeEntries],
            WakeTimerRestoreSnapshotCaptured = WakeTimerRestoreSnapshotCaptured,
            WakeTimerRestoreAcValue = WakeTimerRestoreAcValue,
            WakeTimerRestoreDcValue = WakeTimerRestoreDcValue,
            StandbyConnectivityRestoreSnapshotCaptured = StandbyConnectivityRestoreSnapshotCaptured,
            StandbyConnectivityRestoreAcValue = StandbyConnectivityRestoreAcValue,
            StandbyConnectivityRestoreDcValue = StandbyConnectivityRestoreDcValue,
            DisconnectedStandbyModeRestoreAcValue = DisconnectedStandbyModeRestoreAcValue,
            DisconnectedStandbyModeRestoreDcValue = DisconnectedStandbyModeRestoreDcValue,
            BatteryStandbyHibernateRestoreSnapshotCaptured = BatteryStandbyHibernateRestoreSnapshotCaptured,
            BatteryStandbyHibernateRestoreAcValue = BatteryStandbyHibernateRestoreAcValue,
            BatteryStandbyHibernateRestoreDcValue = BatteryStandbyHibernateRestoreDcValue,
            WiFiDirectAdapterRestoreSnapshotCaptured = WiFiDirectAdapterRestoreSnapshotCaptured,
            WiFiDirectAdapterRestoreInstanceIds = [.. WiFiDirectAdapterRestoreInstanceIds],
            KnownRemoteWakeRequestBackupCaptured = KnownRemoteWakeRequestBackupCaptured,
            KnownRemoteWakeRequestOverrideBackup = new Dictionary<string, int>(KnownRemoteWakeRequestOverrideBackup, StringComparer.OrdinalIgnoreCase),
            ResumeProtectionDelaySeconds = ResumeProtectionDelaySeconds,
            StartMinimized = StartMinimized,
            StartWithWindows = StartWithWindows,
            WindowBoundsCaptured = WindowBoundsCaptured,
            WindowWidth = WindowWidth,
            WindowHeight = WindowHeight,
            WindowX = WindowX,
            WindowY = WindowY,
            LastSuspendUtc = LastSuspendUtc,
            LastResumeUtc = LastResumeUtc,
            LastWakeSummary = LastWakeSummary,
            LastWakeEvidenceSummary = LastWakeEvidenceSummary,
            WakeTimerPolicySummary = WakeTimerPolicySummary,
            StandbyConnectivityPolicySummary = StandbyConnectivityPolicySummary,
            WiFiDirectAdapterPolicySummary = WiFiDirectAdapterPolicySummary,
            BatteryStandbyHibernatePolicySummary = BatteryStandbyHibernatePolicySummary,
            KnownRemoteWakePolicySummary = KnownRemoteWakePolicySummary,
            AutostartPolicySummary = AutostartPolicySummary
        };
    }
}
