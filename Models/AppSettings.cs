namespace SleepSentinel.Models;

public sealed class AppSettings
{
    public PowerPolicyMode PolicyMode { get; set; } = PowerPolicyMode.FollowPowerPlan;
    public ResumeProtectionMode ResumeProtectionMode { get; set; } = ResumeProtectionMode.Hibernate;
    public bool ResumeProtectionEnabled { get; set; } = true;
    public bool ResumeProtectionOnlyForUnattendedWake { get; set; } = true;
    public bool DisableWakeTimers { get; set; }
    public bool DisableStandbyConnectivity { get; set; }
    public bool BlockKnownRemoteWakeRequests { get; set; }
    public bool WakeTimerRestoreSnapshotCaptured { get; set; }
    public int WakeTimerRestoreAcValue { get; set; }
    public int WakeTimerRestoreDcValue { get; set; }
    public bool StandbyConnectivityRestoreSnapshotCaptured { get; set; }
    public int StandbyConnectivityRestoreAcValue { get; set; }
    public int StandbyConnectivityRestoreDcValue { get; set; }
    public int DisconnectedStandbyModeRestoreAcValue { get; set; }
    public int DisconnectedStandbyModeRestoreDcValue { get; set; }
    public bool KnownRemoteWakeRequestBackupCaptured { get; set; }
    public Dictionary<string, int> KnownRemoteWakeRequestOverrideBackup { get; set; } = [];
    public int ResumeProtectionDelaySeconds { get; set; } = 8;
    public bool StartMinimized { get; set; } = true;
    public bool StartWithWindows { get; set; }
    public DateTimeOffset? LastSuspendUtc { get; set; }
    public DateTimeOffset? LastResumeUtc { get; set; }
    public string LastWakeSummary { get; set; } = "无";
    public string WakeTimerPolicySummary { get; set; } = "未检查";
    public string StandbyConnectivityPolicySummary { get; set; } = "未检查";
    public string KnownRemoteWakePolicySummary { get; set; } = "未检查";
}
