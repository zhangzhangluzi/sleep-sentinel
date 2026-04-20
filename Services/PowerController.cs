using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class PowerController : IDisposable
{
    private static readonly TimeSpan ManualResumeSignalWindow = TimeSpan.FromSeconds(20);
    private static readonly TimeSpan StatusSnapshotMaxAge = TimeSpan.FromSeconds(15);
    private const int ResumeAnalysisDelayMilliseconds = 1500;
    private const int DefaultBatteryStandbyHibernateTimeoutSeconds = 600;
    private const int PowerShellTimeoutMilliseconds = 10000;
    private const int WakeTimerDisabledValue = 0;
    private const int WakeTimerEnabledValue = 1;
    private const int StandbyConnectivityDisabledValue = 0;
    private const int StandbyConnectivityEnabledValue = 1;
    private const int StandbyConnectivityManagedByWindowsValue = 2;
    private const int DisconnectedStandbyModeNormalValue = 0;
    private const int DisconnectedStandbyModeAggressiveValue = 1;
    private const int HibernateAfterStandbyDisabledValue = 0;
    private const int HibernateAfterStandbyFallbackRestoreDcValue = int.MaxValue;
    private const int RequestDisplayMask = 1;
    private const int RequestSystemMask = 2;
    private const int RequestAwayModeMask = 4;
    private const int FullRequestOverrideMask = RequestDisplayMask | RequestSystemMask | RequestAwayModeMask;
    private const string PowerRequestOverrideRegistryRoot = @"SYSTEM\CurrentControlSet\Control\Power\PowerRequestOverride";
    private static readonly Regex CurrentAcPowerSettingRegex = new(
        @"(?:Current AC Power Setting Index|当前交流电源设置索引):\s*0x(?<value>[0-9a-fA-F]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CurrentDcPowerSettingRegex = new(
        @"(?:Current DC Power Setting Index|当前直流电源设置索引):\s*0x(?<value>[0-9a-fA-F]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ActiveSchemeRegex = new(
        @"(?<guid>[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ActiveSchemeNameRegex = new(
        @"\((?<name>[^)]+)\)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly string[] ManualWakeIndicators =
    [
        "power button",
        "power switch",
        "电源按钮",
        "电源键",
        "电源开关",
        "sleep button",
        "睡眠按钮",
        "lid",
        "开盖",
        "打开盖",
        "盖子",
        "keyboard",
        "键盘",
        "mouse",
        "鼠标",
        "touchpad",
        "触控板",
        "指纹"
    ];
    private static readonly string[] DeviceWakeIndicators =
    [
        "device",
        "usb",
        "network adapter",
        "wake on lan",
        "pci",
        "蓝牙",
        "bluetooth",
        "网卡"
    ];
    private static readonly string[] TimerWakeIndicators =
    [
        "wake source: timer",
        "wake timer",
        "timer set by",
        "task scheduler",
        "windows update",
        "updateorchestrator",
        "maintenance activator",
        "定时器",
        "计划任务",
        "更新协调器",
        "维护激活器"
    ];
    private static readonly string[] RemoteWakeIndicators =
    [
        "anydesk",
        "awesun",
        "gameviewer",
        "raylink",
        "rustdesk",
        "sunlogin",
        "teamviewer",
        "todesk",
        "uu远控"
    ];
    private static readonly string[] NoWakeHistoryIndicators =
    [
        "history count - 0",
        "wake history count - 0",
        "历史计数 - 0",
        "唤醒历史计数 - 0"
    ];
    private static readonly string[] NoWakeSourceIndicators =
    [
        "wake source count - 0",
        "唤醒源计数 - 0"
    ];
    private static readonly string[] NoActiveWakeTimerIndicators =
    [
        "no active wake timers",
        "没有活动的唤醒定时器",
        "当前没有活动的唤醒定时器"
    ];
    private static readonly string[] DiagnosticsUnavailableIndicators =
    [
        "requires administrator privileges",
        "must be run from an elevated command prompt",
        "administrator privileges",
        "elevated command prompt",
        "此命令需要管理员权限",
        "必须从提升的命令提示符中执行"
    ];
    private readonly SettingsStore _settingsStore;
    private readonly FileLogger _logger;
    private readonly PowerCfgService _powerCfg;
    private readonly WakeDiagnosticsService _wakeDiagnostics;
    private readonly object _stateSync = new();
    private readonly object _manualSignalSync = new();
    private readonly object _resumeProtectionSync = new();
    private readonly object _resumeEvaluationSync = new();
    private readonly object _statusSnapshotSync = new();
    private readonly object _activePowerPlanSync = new();
    private static readonly string WindowsPowerShellPath = ResolveWindowsPowerShellPath();
    private readonly PowerNotificationWindow? _powerNotificationWindow;
    private readonly System.Threading.Timer _resumeTimer;
    private DateTimeOffset? _lastManualResumeSignalUtc;
    private string _lastManualResumeSignalReason = "人工操作";
    private string _activePowerPlanKey = string.Empty;
    private bool _startupWarmupCompleted;
    private bool _startupWarmupInProgress;
    private uint? _resumeProtectionArmedAtTick;
    private int _resumeEvaluationGeneration;
    private int _statusSnapshotRefreshQueued;
    private int _statusSnapshotRevision;
    private bool _disposed;
    private AppSettings _settings;
    private StatusSnapshot _statusSnapshot = StatusSnapshot.Empty;

    private sealed class WiFiDirectAdapterDevice
    {
        public required string Name { get; init; }
        public required string InstanceId { get; init; }
        public bool IsDisabled { get; init; }
        public int ErrorCode { get; init; }
    }

    public PowerController(SettingsStore settingsStore, FileLogger logger, AppSettings settings)
    {
        _settingsStore = settingsStore;
        _logger = logger;
        _settings = settings;
        _powerCfg = new PowerCfgService(logger);
        _wakeDiagnostics = new WakeDiagnosticsService(_powerCfg, logger);
        EnsureSettingsDefaults();
        MigrateLegacyRestoreSnapshotsIfNeeded();
        _resumeTimer = new System.Threading.Timer(
            static state => ((PowerController)state!).OnResumeTimerElapsed(),
            this,
            Timeout.Infinite,
            Timeout.Infinite);
        RememberActivePowerPlan();

        try
        {
            _powerNotificationWindow = new PowerNotificationWindow();
            _powerNotificationWindow.LidStateChanged += OnLidStateChanged;
        }
        catch (Exception ex)
        {
            _logger.Warn($"注册开盖状态监听失败，将退回到键鼠/解锁判断：{ex.Message}");
        }

        SystemEvents.PowerModeChanged += OnPowerModeChanged;
        SystemEvents.SessionSwitch += OnSessionSwitch;
        SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
        ApplyPolicy(_settings.PolicyMode);
        if (_settings.DisableWakeTimers)
        {
            ApplyWakeTimerPolicy();
        }
        else
        {
            _settings.WakeTimerPolicySummary = "未由应用接管；启动时跳过即时核验";
        }
        if (_settings.DisableStandbyConnectivity)
        {
            ApplyStandbyConnectivityPolicy();
        }
        else
        {
            _settings.StandbyConnectivityPolicySummary = "未由应用接管；启动时跳过即时核验";
        }
        if (_settings.DisableWiFiDirectAdapters)
        {
            ApplyWiFiDirectAdapterPolicy();
        }
        else
        {
            _settings.WiFiDirectAdapterPolicySummary = "未由应用接管；启动时跳过即时核验";
        }
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            CaptureBatteryStandbyHibernateRestoreSnapshotIfNeeded();
            ApplyBatteryStandbyHibernatePolicy();
        }
        else
        {
            _settings.BatteryStandbyHibernatePolicySummary = "未由应用接管；启动时跳过即时核验";
        }
        if (_settings.BlockKnownRemoteWakeRequests)
        {
            ApplyKnownRemoteWakePolicy();
        }
        else
        {
            _settings.KnownRemoteWakePolicySummary = "未由应用接管；启动时跳过即时核验";
        }
        _settings.AutostartPolicySummary = _settings.StartWithWindows
            ? "开机自启已启用；启动时跳过即时核验"
            : "开机自启未启用";
        SaveSettingsSnapshot();
        InvalidateStatusSnapshot();
        _logger.Info($"应用启动，当前模式：{DescribeMode(_settings.PolicyMode)}。");
    }

    public event EventHandler? StateChanged;

    public AppSettings CurrentSettings => CloneCurrentSettings();

    public string CurrentStatus
    {
        get
        {
            lock (_stateSync)
            {
                return BuildCurrentStatus(_settings);
            }
        }
    }

    public string CurrentProtectionRuleSummary
    {
        get
        {
            lock (_stateSync)
            {
                return BuildProtectionRuleSummary(_settings);
            }
        }
    }

    public string CurrentCapabilitySummary
    {
        get
        {
            if (StartupWarmupCompleted)
            {
                return GetStatusSnapshot().CapabilitySummary;
            }

            lock (_stateSync)
            {
                return BuildDeferredCapabilitySummary();
            }
        }
    }

    public string CurrentRiskSummary
    {
        get
        {
            if (StartupWarmupCompleted)
            {
                return GetStatusSnapshot().RiskSummary;
            }

            lock (_stateSync)
            {
                return BuildDeferredRiskSummary();
            }
        }
    }

    public string CurrentPowerPlanSummary => GetStatusSnapshot().PowerPlanSummary;

    public string CurrentManagedRemoteEntriesSummary
    {
        get
        {
            lock (_stateSync)
            {
                return BuildManagedRemoteEntriesSummary(_settings);
            }
        }
    }

    public string CurrentWakeTimerQuickState => GetStatusSnapshot().WakeTimerQuickState;

    public string CurrentStandbyConnectivityQuickState => GetStatusSnapshot().StandbyConnectivityQuickState;

    public string CurrentWiFiDirectQuickState => GetStatusSnapshot().WiFiDirectQuickState;

    public string CurrentBatteryStandbyHibernateQuickState => GetStatusSnapshot().BatteryStandbyHibernateQuickState;

    public string CurrentRemoteWakeQuickState => GetStatusSnapshot().RemoteWakeQuickState;

    public bool StartupWarmupCompleted
    {
        get
        {
            lock (_stateSync)
            {
                return _startupWarmupCompleted;
            }
        }
    }

    private StatusSnapshot GetStatusSnapshot()
    {
        StatusSnapshot snapshot;
        lock (_statusSnapshotSync)
        {
            snapshot = _statusSnapshot;
        }

        if (snapshot == StatusSnapshot.Empty)
        {
            InvalidateStatusSnapshot();
            lock (_statusSnapshotSync)
            {
                snapshot = _statusSnapshot;
            }
        }

        if (snapshot != StatusSnapshot.Empty
            && DateTimeOffset.UtcNow - snapshot.GeneratedAtUtc > StatusSnapshotMaxAge
            && StartupWarmupCompleted)
        {
            QueueStatusSnapshotRefresh();
        }

        return snapshot;
    }

    private void InvalidateStatusSnapshot()
    {
        lock (_stateSync)
        {
            _statusSnapshotRevision++;
            lock (_statusSnapshotSync)
            {
                _statusSnapshot = BuildDeferredStatusSnapshot();
            }
        }
    }

    public void SetPolicyMode(PowerPolicyMode mode)
    {
        lock (_stateSync)
        {
            if (_settings.PolicyMode == mode)
            {
                return;
            }

            CancelPendingResumeProtection();
            CancelPendingResumeEvaluation();
            _settings.PolicyMode = mode;
            EnsureSettingsDefaults();
            SaveSettingsSnapshot();
            ApplyPolicy(_settings.PolicyMode);
            InvalidateStatusSnapshot();
            _logger.Info($"设置已更新：模式={DescribeMode(_settings.PolicyMode)}。");
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void UpdateSettings(AppSettings updatedSettings)
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            CancelPendingResumeProtection();
            CancelPendingResumeEvaluation();
            var previousSettings = _settings;
            var previousManagedRemoteEntries = RemoteWakeBlockCatalog.GetManagedEntries(previousSettings.CustomRemoteWakeEntries);
            _settings = updatedSettings;
            EnsureSettingsDefaults();
            SaveSettingsSnapshot();
            ApplyPolicy(_settings.PolicyMode);
            if (_settings.DisableWakeTimers)
            {
                if (!previousSettings.DisableWakeTimers)
                {
                    CaptureWakeTimerRestoreSnapshotIfNeeded();
                }

                ApplyWakeTimerPolicy();
            }
            else if (previousSettings.DisableWakeTimers)
            {
                RestoreWakeTimerPolicy();
            }
            else
            {
                RefreshWakeTimerPolicySummary(persistSnapshot: false);
            }
            if (_settings.DisableStandbyConnectivity)
            {
                if (!previousSettings.DisableStandbyConnectivity)
                {
                    CaptureStandbyConnectivityRestoreSnapshotIfNeeded();
                }

                ApplyStandbyConnectivityPolicy();
            }
            else if (previousSettings.DisableStandbyConnectivity)
            {
                RestoreStandbyConnectivityPolicy();
            }
            else
            {
                RefreshStandbyConnectivityPolicySummary(persistSnapshot: false);
            }
            if (_settings.DisableWiFiDirectAdapters)
            {
                if (!previousSettings.DisableWiFiDirectAdapters)
                {
                    CaptureWiFiDirectAdapterRestoreSnapshotIfNeeded();
                }

                ApplyWiFiDirectAdapterPolicy();
            }
            else if (previousSettings.DisableWiFiDirectAdapters)
            {
                RestoreWiFiDirectAdapterPolicy();
            }
            else
            {
                RefreshWiFiDirectAdapterPolicySummary(persistSnapshot: false);
            }
            if (_settings.EnforceBatteryStandbyHibernate)
            {
                if (!previousSettings.EnforceBatteryStandbyHibernate)
                {
                    CaptureBatteryStandbyHibernateRestoreSnapshotIfNeeded();
                }

                ApplyBatteryStandbyHibernatePolicy();
            }
            else if (previousSettings.EnforceBatteryStandbyHibernate)
            {
                RestoreBatteryStandbyHibernatePolicy();
            }
            else
            {
                RefreshBatteryStandbyHibernatePolicySummary(persistSnapshot: false);
            }
            if (_settings.BlockKnownRemoteWakeRequests)
            {
                var currentManagedRemoteEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);
                if (!previousSettings.BlockKnownRemoteWakeRequests)
                {
                    CaptureKnownRemoteWakeRestoreSnapshotIfNeeded(currentManagedRemoteEntries);
                }
                else
                {
                    ReconcileKnownRemoteWakeManagedEntries(previousManagedRemoteEntries, currentManagedRemoteEntries);
                }

                ApplyKnownRemoteWakePolicy(currentManagedRemoteEntries);
            }
            else if (previousSettings.BlockKnownRemoteWakeRequests)
            {
                RestoreKnownRemoteWakePolicy(previousManagedRemoteEntries);
            }
            else
            {
                RefreshKnownRemoteWakePolicySummary(persistSnapshot: false);
            }
            EnsureAutostartMatchesSettings();
            SaveSettingsSnapshot();
            InvalidateStatusSnapshot();
            _logger.Info($"设置已更新：模式={DescribeMode(_settings.PolicyMode)}，恢复保护={_settings.ResumeProtectionEnabled}，待机联网拦截={_settings.DisableStandbyConnectivity}，Wi-Fi Direct 禁用={_settings.DisableWiFiDirectAdapters}，电池兜底休眠={_settings.EnforceBatteryStandbyHibernate}（{DescribePowerSettingDuration(GetBatteryStandbyHibernateTimeoutSeconds())}），远控拦截={_settings.BlockKnownRemoteWakeRequests}，自定义远控={_settings.CustomRemoteWakeEntries.Count} 条。");
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void SleepNow()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            CancelPendingResumeProtection($"用户手动操作，已取消待执行的自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
            RequestSuspend(PowerState.Suspend, "用户手动请求立即睡眠。");
        }
    }

    public void HibernateNow()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            CancelPendingResumeProtection($"用户手动操作，已取消待执行的自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
            RequestSuspend(PowerState.Hibernate, "用户手动请求立即休眠。");
        }
    }

    public string CollectWakeDiagnostics()
    {
        var snapshot = CollectWakeDiagnosticSnapshot(includePowerRequests: false, includeSleepStudy: false);
        return BuildWakeDiagnosticsText(snapshot, includePowerRequests: false, includeSleepStudy: false);
    }

    public string CollectPowerRequestDiagnostics()
    {
        var snapshot = CollectWakeDiagnosticSnapshot(includePowerRequests: true, includeSleepStudy: false);
        return $"requests:{Environment.NewLine}{snapshot.PowerRequestsText}{Environment.NewLine}{Environment.NewLine}requestsoverride:{Environment.NewLine}{snapshot.RequestOverridesText}";
    }

    public WakeDiagnosticSnapshot CollectWakeDiagnosticSnapshot(bool includePowerRequests = true, bool includeSleepStudy = true)
    {
        var settingsSnapshot = CloneCurrentSettings();
        return _wakeDiagnostics.CollectSnapshot(
            settingsSnapshot.LastSuspendUtc?.AddMinutes(-10),
            includePowerRequests,
            includeSleepStudy,
            settingsSnapshot.CustomRemoteWakeEntries);
    }

    public string FormatWakeDiagnosticSnapshot(WakeDiagnosticSnapshot snapshot, bool includePowerRequests = true, bool includeSleepStudy = true)
    {
        return BuildWakeDiagnosticsText(snapshot, includePowerRequests, includeSleepStudy);
    }

    public string CollectFullWakeDiagnosticsText()
    {
        var snapshot = CollectWakeDiagnosticSnapshot(includePowerRequests: true, includeSleepStudy: true);
        return BuildWakeDiagnosticsText(snapshot, includePowerRequests: true, includeSleepStudy: true);
    }

    public IReadOnlyList<string> SuggestCustomRemoteWakeEntries()
    {
        return _wakeDiagnostics.SuggestCustomRemoteWakeEntries(CloneCurrentSettings().CustomRemoteWakeEntries);
    }

    public void UpdateCustomRemoteWakeEntries(IEnumerable<string> rawEntries)
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            var normalized = RemoteWakeBlockCatalog.NormalizeCustomEntries(rawEntries);
            if (_settings.CustomRemoteWakeEntries.SequenceEqual(normalized, StringComparer.OrdinalIgnoreCase))
            {
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.CustomRemoteWakeEntries = normalized.ToList();
            _logger.Info(normalized.Count == 0
                ? "已清空自定义远控拦截名单。"
                : $"已更新自定义远控拦截名单，共 {normalized.Count} 条。");
            UpdateSettings(updatedSettings);
        }
    }

    public void ReapplyAllManagedSettings()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            CancelPendingResumeProtection();
            ApplyPolicy(_settings.PolicyMode);

            if (_settings.DisableWakeTimers)
            {
                ApplyWakeTimerPolicy();
            }
            else
            {
                RefreshWakeTimerPolicySummary(persistSnapshot: false);
            }

            if (_settings.DisableStandbyConnectivity)
            {
                ApplyStandbyConnectivityPolicy();
            }
            else
            {
                RefreshStandbyConnectivityPolicySummary(persistSnapshot: false);
            }

            if (_settings.DisableWiFiDirectAdapters)
            {
                ApplyWiFiDirectAdapterPolicy();
            }
            else
            {
                RefreshWiFiDirectAdapterPolicySummary(persistSnapshot: false);
            }

            if (_settings.EnforceBatteryStandbyHibernate)
            {
                ApplyBatteryStandbyHibernatePolicy();
            }
            else
            {
                RefreshBatteryStandbyHibernatePolicySummary(persistSnapshot: false);
            }

            if (_settings.BlockKnownRemoteWakeRequests)
            {
                ApplyKnownRemoteWakePolicy();
            }
            else
            {
                RefreshKnownRemoteWakePolicySummary(persistSnapshot: false);
            }

            EnsureAutostartMatchesSettings();
            SaveSettingsSnapshot();
            _logger.Info("已重新应用当前全部设置。");
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void CompleteDeferredStartupValidation()
    {
        var shouldNotify = false;

        lock (_stateSync)
        {
            if (_startupWarmupCompleted || _startupWarmupInProgress)
            {
                return;
            }

            _startupWarmupInProgress = true;
        }

        try
        {
            _logger.Info("正在补全启动后的延迟状态校验。");

            if (!TryRunDeferredStartupRefresh(static controller => controller.RefreshWakeTimerPolicySummary(persistSnapshot: false))
                || !TryRunDeferredStartupRefresh(static controller => controller.RefreshStandbyConnectivityPolicySummary(persistSnapshot: false))
                || !TryRunDeferredStartupRefresh(static controller => controller.RefreshWiFiDirectAdapterPolicySummary(persistSnapshot: false))
                || !TryRunDeferredStartupRefresh(static controller => controller.RefreshBatteryStandbyHibernatePolicySummary(persistSnapshot: false))
                || !TryRunDeferredStartupRefresh(static controller => controller.RefreshKnownRemoteWakePolicySummary(persistSnapshot: false))
                || !TryRunDeferredStartupRefresh(static controller => controller.RefreshAutostartPolicySummary()))
            {
                return;
            }

            lock (_stateSync)
            {
                if (_disposed)
                {
                    return;
                }

                _startupWarmupCompleted = true;
                shouldNotify = true;
                InvalidateStatusSnapshot();
                SaveSettingsSnapshot();
            }
        }
        finally
        {
            lock (_stateSync)
            {
                _startupWarmupInProgress = false;
            }
        }

        if (shouldNotify)
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void ReapplyWakeTimerPolicy()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.DisableWakeTimers)
            {
                ApplyWakeTimerPolicy();
            }
            else
            {
                RefreshWakeTimerPolicySummary(persistSnapshot: false);
                _logger.Info(_settings.WakeTimerPolicySummary);
            }

            SaveSettingsSnapshot();
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void ReapplyKnownRemoteWakePolicy()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.BlockKnownRemoteWakeRequests)
            {
                ApplyKnownRemoteWakePolicy();
            }
            else
            {
                RefreshKnownRemoteWakePolicySummary(persistSnapshot: false);
                SaveSettingsSnapshot();
            }

            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void ReapplyStandbyConnectivityPolicy()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.DisableStandbyConnectivity)
            {
                ApplyStandbyConnectivityPolicy();
            }
            else
            {
                RefreshStandbyConnectivityPolicySummary(persistSnapshot: false);
                SaveSettingsSnapshot();
                _logger.Info(_settings.StandbyConnectivityPolicySummary);
            }

            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void ReapplyWiFiDirectAdapterPolicy()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.DisableWiFiDirectAdapters)
            {
                ApplyWiFiDirectAdapterPolicy();
            }
            else
            {
                RefreshWiFiDirectAdapterPolicySummary(persistSnapshot: false);
                SaveSettingsSnapshot();
                _logger.Info(_settings.WiFiDirectAdapterPolicySummary);
            }

            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void ReapplyBatteryStandbyHibernatePolicy()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.EnforceBatteryStandbyHibernate)
            {
                ApplyBatteryStandbyHibernatePolicy();
            }
            else
            {
                RefreshBatteryStandbyHibernatePolicySummary(persistSnapshot: false);
                SaveSettingsSnapshot();
                _logger.Info(_settings.BatteryStandbyHibernatePolicySummary);
            }

            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void BlockSoftwareWake()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.DisableWakeTimers)
            {
                _logger.Info("软件/定时器唤醒拦截已启用，正在重新应用当前电源计划策略。");
                ApplyWakeTimerPolicy();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.DisableWakeTimers = true;
            _logger.Info("已通过快捷按钮启用软件/定时器唤醒拦截。");
            UpdateSettings(updatedSettings);
        }
    }

    public void RestoreSoftwareWake()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (!_settings.DisableWakeTimers)
            {
                _logger.Info("软件/定时器唤醒拦截当前未启用，无需恢复。");
                RefreshWakeTimerPolicySummary();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.DisableWakeTimers = false;
            _logger.Info("已通过快捷按钮请求恢复软件/定时器唤醒策略。");
            UpdateSettings(updatedSettings);
        }
    }

    public void BlockStandbyConnectivityWake()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.DisableStandbyConnectivity)
            {
                _logger.Info("待机联网拦截已启用，正在重新应用当前电源计划策略。");
                ApplyStandbyConnectivityPolicy();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.DisableStandbyConnectivity = true;
            _logger.Info("已通过快捷按钮启用待机联网拦截。");
            UpdateSettings(updatedSettings);
        }
    }

    public void RestoreStandbyConnectivityWake()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (!_settings.DisableStandbyConnectivity)
            {
                _logger.Info("待机联网拦截当前未启用，无需恢复。");
                RefreshStandbyConnectivityPolicySummary();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.DisableStandbyConnectivity = false;
            _logger.Info("已通过快捷按钮请求恢复待机联网策略。");
            UpdateSettings(updatedSettings);
        }
    }

    public void DisableWiFiDirectAdapters()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.DisableWiFiDirectAdapters)
            {
                _logger.Info("Wi-Fi Direct 虚拟适配器禁用策略已启用，正在重新应用。");
                ApplyWiFiDirectAdapterPolicy();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.DisableWiFiDirectAdapters = true;
            _logger.Info("已通过快捷按钮启用 Wi-Fi Direct 虚拟适配器禁用策略。");
            UpdateSettings(updatedSettings);
        }
    }

    public void RestoreWiFiDirectAdapters()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (!_settings.DisableWiFiDirectAdapters)
            {
                _logger.Info("Wi-Fi Direct 虚拟适配器禁用策略当前未启用，无需恢复。");
                RefreshWiFiDirectAdapterPolicySummary();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.DisableWiFiDirectAdapters = false;
            _logger.Info("已通过快捷按钮请求恢复 Wi-Fi Direct 虚拟适配器状态。");
            UpdateSettings(updatedSettings);
        }
    }

    public void EnableBatteryStandbyHibernateFallback()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.EnforceBatteryStandbyHibernate)
            {
                _logger.Info("电池兜底休眠已启用，正在重新应用当前电源计划策略。");
                ApplyBatteryStandbyHibernatePolicy();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.EnforceBatteryStandbyHibernate = true;
            _logger.Info("已通过快捷按钮启用电池待机兜底休眠。");
            UpdateSettings(updatedSettings);
        }
    }

    public void RestoreBatteryStandbyHibernateFallback()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (!_settings.EnforceBatteryStandbyHibernate)
            {
                _logger.Info("电池兜底休眠当前未启用，无需恢复。");
                RefreshBatteryStandbyHibernatePolicySummary();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.EnforceBatteryStandbyHibernate = false;
            _logger.Info("已通过快捷按钮请求恢复电池待机休眠策略。");
            UpdateSettings(updatedSettings);
        }
    }

    public void BlockKnownRemoteWakeRequests()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (_settings.BlockKnownRemoteWakeRequests)
            {
                _logger.Info("常见远程软件保持唤醒拦截已启用，正在重新应用规则。");
                ApplyKnownRemoteWakePolicy();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.BlockKnownRemoteWakeRequests = true;
            _logger.Info("已通过快捷按钮启用常见远程软件保持唤醒拦截。");
            UpdateSettings(updatedSettings);
        }
    }

    public void RestoreKnownRemoteWakeRequests()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            if (!_settings.BlockKnownRemoteWakeRequests)
            {
                _logger.Info("常见远程软件保持唤醒拦截当前未启用，无需恢复。");
                RefreshKnownRemoteWakePolicySummary();
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            var updatedSettings = CloneCurrentSettings();
            updatedSettings.BlockKnownRemoteWakeRequests = false;
            _logger.Info("已通过快捷按钮请求恢复常见远程软件保持唤醒策略。");
            UpdateSettings(updatedSettings);
        }
    }

    public void Dispose()
    {
        _disposed = true;
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        SystemEvents.SessionSwitch -= OnSessionSwitch;
        SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
        if (_powerNotificationWindow is not null)
        {
            _powerNotificationWindow.LidStateChanged -= OnLidStateChanged;
            _powerNotificationWindow.Dispose();
        }
        CancelPendingResumeEvaluation();
        CancelPendingResumeProtection();
        _resumeTimer.Dispose();
        NativeMethods.SetThreadExecutionState(ExecutionState.EsContinuous);
    }

    private void ApplyPolicy(PowerPolicyMode mode)
    {
        var executionState = mode == PowerPolicyMode.KeepAwakeIndefinitely
            ? ExecutionState.EsContinuous | ExecutionState.EsSystemRequired
            : ExecutionState.EsContinuous;

        if (NativeMethods.SetThreadExecutionState(executionState) == 0)
        {
            _logger.Error($"设置执行状态失败：{DescribeMode(mode)}。");
            return;
        }

        if (mode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            _logger.Info("已启用无限保持唤醒。");
        }
        else
        {
            _logger.Info("已切换为遵循电源计划。");
        }
    }

    private void ApplyWakeTimerPolicy()
    {
        var failures = SetWakeTimerPolicy(WakeTimerDisabledValue, WakeTimerDisabledValue);

        if (failures.Length == 0)
        {
            RefreshWakeTimerPolicySummary();
            _logger.Info(_settings.WakeTimerPolicySummary);
            return;
        }

        RefreshWakeTimerPolicySummary();
        _logger.Warn($"切换唤醒定时器策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RestoreWakeTimerPolicy()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var hadSnapshot = snapshot.WakeTimerCaptured;
        var acValue = hadSnapshot ? snapshot.WakeTimerAcValue : WakeTimerEnabledValue;
        var dcValue = hadSnapshot ? snapshot.WakeTimerDcValue : WakeTimerEnabledValue;
        var failures = SetWakeTimerPolicy(acValue, dcValue);

        if (failures.Length == 0)
        {
            snapshot.WakeTimerCaptured = false;
            snapshot.WakeTimerAcValue = 0;
            snapshot.WakeTimerDcValue = 0;
            RefreshWakeTimerPolicySummary();
            _logger.Info(_settings.WakeTimerPolicySummary);
            return;
        }

        RefreshWakeTimerPolicySummary();
        _logger.Warn($"恢复唤醒定时器策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshWakeTimerPolicySummary(bool persistSnapshot = true)
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var status = EvaluateWakeTimerProtection();
        if (_settings.DisableWakeTimers)
        {
            _settings.WakeTimerPolicySummary = status.Kind switch
            {
                ProtectionStatusKind.Applied => snapshot.WakeTimerCaptured
                    ? "当前电源计划的唤醒定时器已禁用（AC/DC，可恢复到原始设置）"
                    : "当前电源计划的唤醒定时器已禁用（AC/DC，历史基线缺失时将回退到常规启用）",
                ProtectionStatusKind.Partial => "已配置为由应用禁用当前电源计划的唤醒定时器，但当前仅部分生效",
                ProtectionStatusKind.Unknown => "已配置为由应用禁用当前电源计划的唤醒定时器，但暂时无法验证当前状态",
                _ => "已配置为由应用禁用当前电源计划的唤醒定时器，但当前未生效，请检查权限或日志"
            };
        }
        else if (TryGetCurrentWakeTimerIndices(out var acValue, out var dcValue))
        {
            _settings.WakeTimerPolicySummary = acValue == WakeTimerDisabledValue && dcValue == WakeTimerDisabledValue
                ? "未由应用接管；系统当前已禁用唤醒定时器（AC/DC）"
                : $"未由应用接管；系统当前唤醒定时器：AC={acValue}，DC={dcValue}";
        }
        else
        {
            _settings.WakeTimerPolicySummary = "未由应用接管；保留系统当前唤醒定时器策略";
        }

        if (persistSnapshot)
        {
            SaveSettingsSnapshot();
        }
    }

    private void CaptureWakeTimerRestoreSnapshotIfNeeded()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        if (snapshot.WakeTimerCaptured)
        {
            return;
        }

        if (!TryGetCurrentWakeTimerIndices(out var acValue, out var dcValue))
        {
            _logger.Warn("未能记录唤醒定时器原始 AC/DC 设置；恢复时将回退到常规启用。");
            return;
        }

        snapshot.WakeTimerAcValue = acValue;
        snapshot.WakeTimerDcValue = dcValue;
        snapshot.WakeTimerCaptured = true;
        SaveSettingsSnapshot();
        _logger.Info($"已记录唤醒定时器原始设置：AC={acValue}，DC={dcValue}。");
    }

    private void ApplyStandbyConnectivityPolicy()
    {
        var failures = SetStandbyConnectivityPolicy(
            StandbyConnectivityDisabledValue,
            StandbyConnectivityDisabledValue,
            DisconnectedStandbyModeAggressiveValue,
            DisconnectedStandbyModeAggressiveValue);

        if (failures.Length == 0)
        {
            RefreshStandbyConnectivityPolicySummary();
            _logger.Info(_settings.StandbyConnectivityPolicySummary);
            return;
        }

        RefreshStandbyConnectivityPolicySummary();
        _logger.Warn($"切换待机联网策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RestoreStandbyConnectivityPolicy()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var hadSnapshot = snapshot.StandbyConnectivityCaptured;
        var connectivityAcValue = hadSnapshot ? snapshot.StandbyConnectivityAcValue : StandbyConnectivityManagedByWindowsValue;
        var connectivityDcValue = hadSnapshot ? snapshot.StandbyConnectivityDcValue : StandbyConnectivityManagedByWindowsValue;
        var disconnectedAcValue = hadSnapshot ? snapshot.DisconnectedStandbyModeAcValue : DisconnectedStandbyModeNormalValue;
        var disconnectedDcValue = hadSnapshot ? snapshot.DisconnectedStandbyModeDcValue : DisconnectedStandbyModeNormalValue;
        var failures = SetStandbyConnectivityPolicy(
            connectivityAcValue,
            connectivityDcValue,
            disconnectedAcValue,
            disconnectedDcValue);

        if (failures.Length == 0)
        {
            snapshot.StandbyConnectivityCaptured = false;
            snapshot.StandbyConnectivityAcValue = 0;
            snapshot.StandbyConnectivityDcValue = 0;
            snapshot.DisconnectedStandbyModeAcValue = 0;
            snapshot.DisconnectedStandbyModeDcValue = 0;
            RefreshStandbyConnectivityPolicySummary();
            _logger.Info(_settings.StandbyConnectivityPolicySummary);
            return;
        }

        RefreshStandbyConnectivityPolicySummary();
        _logger.Warn($"恢复待机联网策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshStandbyConnectivityPolicySummary(bool persistSnapshot = true)
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        if (_settings.DisableStandbyConnectivity)
        {
            var status = EvaluateStandbyConnectivityProtection();
            _settings.StandbyConnectivityPolicySummary = status.Kind switch
            {
                ProtectionStatusKind.Applied => snapshot.StandbyConnectivityCaptured
                    ? "当前电源计划已关闭待机联网（AC/DC），并启用主动断网待机模式，可恢复到原始设置"
                    : "当前电源计划已关闭待机联网（AC/DC），并启用主动断网待机模式；历史基线缺失时将回退到 Windows 管理/正常模式",
                ProtectionStatusKind.Partial => "已配置为关闭待机联网并启用主动断网待机模式，但当前仅部分生效",
                ProtectionStatusKind.Unknown => "已配置为关闭待机联网并启用主动断网待机模式，但暂时无法验证当前状态",
                _ => "已配置为关闭待机联网并启用主动断网待机模式，但当前未生效，请检查权限或日志"
            };
            if (persistSnapshot)
            {
                SaveSettingsSnapshot();
            }
            return;
        }

        if (TryGetCurrentStandbyConnectivityIndices(
                out var connectivityAcValue,
                out var connectivityDcValue,
                out var disconnectedAcValue,
                out var disconnectedDcValue))
        {
            _settings.StandbyConnectivityPolicySummary =
                $"未由应用接管；系统当前待机联网：AC={DescribeStandbyConnectivityValue(connectivityAcValue)}，DC={DescribeStandbyConnectivityValue(connectivityDcValue)}；断网待机模式：AC={DescribeDisconnectedStandbyModeValue(disconnectedAcValue)}，DC={DescribeDisconnectedStandbyModeValue(disconnectedDcValue)}";
        }
        else
        {
            _settings.StandbyConnectivityPolicySummary = "未由应用接管；保留系统当前待机联网策略";
        }

        if (persistSnapshot)
        {
            SaveSettingsSnapshot();
        }
    }

    private void CaptureStandbyConnectivityRestoreSnapshotIfNeeded()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        if (snapshot.StandbyConnectivityCaptured)
        {
            return;
        }

        if (!TryGetCurrentStandbyConnectivityIndices(
                out var connectivityAcValue,
                out var connectivityDcValue,
                out var disconnectedAcValue,
                out var disconnectedDcValue))
        {
            _logger.Warn("未能记录待机联网原始 AC/DC 设置；恢复时将回退到 Windows 管理/正常模式。");
            return;
        }

        snapshot.StandbyConnectivityAcValue = connectivityAcValue;
        snapshot.StandbyConnectivityDcValue = connectivityDcValue;
        snapshot.DisconnectedStandbyModeAcValue = disconnectedAcValue;
        snapshot.DisconnectedStandbyModeDcValue = disconnectedDcValue;
        snapshot.StandbyConnectivityCaptured = true;
        SaveSettingsSnapshot();
        _logger.Info($"已记录待机联网原始设置：联网 AC={connectivityAcValue}，DC={connectivityDcValue}；断网待机 AC={disconnectedAcValue}，DC={disconnectedDcValue}。");
    }

    private void ApplyWiFiDirectAdapterPolicy()
    {
        if (!TryGetWiFiDirectAdapters(out var adapters, out var failureMessage))
        {
            _settings.WiFiDirectAdapterPolicySummary = "查询 Wi-Fi Direct 虚拟适配器状态失败，请检查日志";
            SaveSettingsSnapshot();
            _logger.Warn($"查询 Wi-Fi Direct 虚拟适配器状态失败：{failureMessage}");
            return;
        }

        if (adapters.Count == 0)
        {
            _settings.WiFiDirectAdapterPolicySummary = "未检测到 Wi-Fi Direct 虚拟适配器";
            SaveSettingsSnapshot();
            _logger.Info(_settings.WiFiDirectAdapterPolicySummary);
            return;
        }

        var failures = new List<string>();
        foreach (var adapter in adapters.Where(static adapter => !adapter.IsDisabled))
        {
            var result = SetWiFiDirectAdapterEnabled(adapter.InstanceId, enable: false);
            if (!string.IsNullOrWhiteSpace(result))
            {
                failures.Add($"{adapter.Name}: {result}");
            }
        }

        var effectiveAdapters = adapters;
        if (TryGetWiFiDirectAdapters(out var refreshedAdapters, out _))
        {
            effectiveAdapters = refreshedAdapters;
        }

        var disabledCount = effectiveAdapters.Count(static adapter => adapter.IsDisabled);
        var allDisabled = effectiveAdapters.Count > 0 && disabledCount == effectiveAdapters.Count;
        if (allDisabled)
        {
            _settings.WiFiDirectAdapterPolicySummary = _settings.WiFiDirectAdapterRestoreSnapshotCaptured
                ? $"已禁用当前 Wi-Fi Direct 虚拟适配器（当前检测到 {effectiveAdapters.Count} 个，均已禁用，可恢复到系统管理状态）"
                : $"已禁用当前 Wi-Fi Direct 虚拟适配器（当前检测到 {effectiveAdapters.Count} 个，均已禁用，缺少历史基线时将按保持禁用处理）";
            SaveSettingsSnapshot();
            _logger.Info(_settings.WiFiDirectAdapterPolicySummary);
            return;
        }

        _settings.WiFiDirectAdapterPolicySummary = $"应用 Wi-Fi Direct 虚拟适配器禁用策略时收到输出，请检查日志（当前 {disabledCount}/{effectiveAdapters.Count} 个已禁用）";
        SaveSettingsSnapshot();
        _logger.Warn($"应用 Wi-Fi Direct 虚拟适配器禁用策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RestoreWiFiDirectAdapterPolicy()
    {
        var hadSnapshot = _settings.WiFiDirectAdapterRestoreSnapshotCaptured;
        if (!TryGetWiFiDirectAdapters(out var adapters, out var failureMessage))
        {
            _settings.WiFiDirectAdapterPolicySummary = "查询 Wi-Fi Direct 虚拟适配器状态失败，请检查日志";
            SaveSettingsSnapshot();
            _logger.Warn($"查询 Wi-Fi Direct 虚拟适配器状态失败：{failureMessage}");
            return;
        }

        if (adapters.Count == 0)
        {
            _settings.WiFiDirectAdapterPolicySummary = "未检测到 Wi-Fi Direct 虚拟适配器";
            _settings.WiFiDirectAdapterRestoreSnapshotCaptured = false;
            _settings.WiFiDirectAdapterRestoreInstanceIds = [];
            SaveSettingsSnapshot();
            _logger.Info(_settings.WiFiDirectAdapterPolicySummary);
            return;
        }

        var failures = new List<string>();
        if (hadSnapshot)
        {
            var adaptersByInstanceId = adapters.ToDictionary(static adapter => adapter.InstanceId, StringComparer.OrdinalIgnoreCase);
            foreach (var instanceId in _settings.WiFiDirectAdapterRestoreInstanceIds.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!adaptersByInstanceId.TryGetValue(instanceId, out var adapter) || !adapter.IsDisabled)
                {
                    continue;
                }

                var result = SetWiFiDirectAdapterEnabled(instanceId, enable: true);
                if (!string.IsNullOrWhiteSpace(result))
                {
                    failures.Add($"{adapter.Name}: {result}");
                }
            }
        }

        var effectiveAdapters = adapters;
        if (TryGetWiFiDirectAdapters(out var refreshedAdapters, out _))
        {
            effectiveAdapters = refreshedAdapters;
        }

        var disabledCount = effectiveAdapters.Count(static adapter => adapter.IsDisabled);
        var enabledCount = effectiveAdapters.Count - disabledCount;
        var allEnabled = effectiveAdapters.Count > 0 && enabledCount == effectiveAdapters.Count;
        if (allEnabled)
        {
            _settings.WiFiDirectAdapterPolicySummary = hadSnapshot
                ? "已恢复当前 Wi-Fi Direct 虚拟适配器到系统管理状态（次级适配器可能由 Windows 按需重建）"
                : "已停止接管 Wi-Fi Direct 虚拟适配器；当前检测到的适配器均保持启用";
            _settings.WiFiDirectAdapterRestoreSnapshotCaptured = false;
            _settings.WiFiDirectAdapterRestoreInstanceIds = [];
            SaveSettingsSnapshot();
            _logger.Info(_settings.WiFiDirectAdapterPolicySummary);
            return;
        }

        _settings.WiFiDirectAdapterPolicySummary = "尝试恢复 Wi-Fi Direct 虚拟适配器状态时收到输出，请检查日志";
        SaveSettingsSnapshot();
        _logger.Warn($"恢复 Wi-Fi Direct 虚拟适配器状态时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshWiFiDirectAdapterPolicySummary(bool persistSnapshot = true)
    {
        if (!TryGetWiFiDirectAdapters(out var adapters, out var failureMessage))
        {
            _settings.WiFiDirectAdapterPolicySummary = "未能查询 Wi-Fi Direct 虚拟适配器状态，请检查日志";
            if (persistSnapshot)
            {
                SaveSettingsSnapshot();
            }
            _logger.Warn($"刷新 Wi-Fi Direct 虚拟适配器摘要失败：{failureMessage}");
            return;
        }

        if (adapters.Count == 0)
        {
            _settings.WiFiDirectAdapterPolicySummary = "未检测到 Wi-Fi Direct 虚拟适配器";
            if (persistSnapshot)
            {
                SaveSettingsSnapshot();
            }
            return;
        }

        var disabledCount = adapters.Count(static adapter => adapter.IsDisabled);
        if (_settings.DisableWiFiDirectAdapters)
        {
            _settings.WiFiDirectAdapterPolicySummary = disabledCount == adapters.Count
                ? (_settings.WiFiDirectAdapterRestoreSnapshotCaptured
                    ? $"已禁用当前 Wi-Fi Direct 虚拟适配器（当前检测到 {adapters.Count} 个，均已禁用，可恢复到系统管理状态）"
                    : $"已禁用当前 Wi-Fi Direct 虚拟适配器（当前检测到 {adapters.Count} 个，均已禁用，缺少历史基线时将按保持禁用处理）")
                : disabledCount > 0
                    ? $"已配置为禁用 Wi-Fi Direct 虚拟适配器，但当前仅 {disabledCount}/{adapters.Count} 个已禁用"
                    : "已配置为禁用 Wi-Fi Direct 虚拟适配器，但当前未生效，请检查权限或日志";
        }
        else
        {
            _settings.WiFiDirectAdapterPolicySummary = disabledCount == 0
                ? "未由应用接管；Wi-Fi Direct 虚拟适配器当前保持启用"
                : $"未由应用接管；系统当前已有 {disabledCount}/{adapters.Count} 个 Wi-Fi Direct 虚拟适配器处于禁用";
        }

        if (persistSnapshot)
        {
            SaveSettingsSnapshot();
        }
    }

    private void CaptureWiFiDirectAdapterRestoreSnapshotIfNeeded()
    {
        if (_settings.WiFiDirectAdapterRestoreSnapshotCaptured)
        {
            return;
        }

        if (!TryGetWiFiDirectAdapters(out var adapters, out var failureMessage))
        {
            _logger.Warn($"未能记录 Wi-Fi Direct 虚拟适配器恢复基线：{failureMessage}");
            return;
        }

        _settings.WiFiDirectAdapterRestoreInstanceIds = adapters
            .Where(static adapter => !adapter.IsDisabled)
            .Select(static adapter => adapter.InstanceId)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        _settings.WiFiDirectAdapterRestoreSnapshotCaptured = true;
        SaveSettingsSnapshot();
        _logger.Info($"已记录 Wi-Fi Direct 虚拟适配器恢复基线（待恢复 {_settings.WiFiDirectAdapterRestoreInstanceIds.Count} 个）。");
    }

    private void ApplyBatteryStandbyHibernatePolicy()
    {
        var timeoutSeconds = GetBatteryStandbyHibernateTimeoutSeconds();
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var acValue = snapshot.BatteryStandbyHibernateCaptured
            ? snapshot.BatteryStandbyHibernateAcValue
            : TryGetCurrentHibernateAfterStandbyIndices(out var currentAcValue, out _)
                ? currentAcValue
                : HibernateAfterStandbyDisabledValue;
        var failures = SetHibernateAfterStandbyPolicy(acValue, timeoutSeconds);

        if (failures.Length == 0)
        {
            RefreshBatteryStandbyHibernatePolicySummary();
            _logger.Info(_settings.BatteryStandbyHibernatePolicySummary);
            return;
        }

        RefreshBatteryStandbyHibernatePolicySummary();
        _logger.Warn($"切换电池待机休眠策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RestoreBatteryStandbyHibernatePolicy()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var hadSnapshot = snapshot.BatteryStandbyHibernateCaptured;
        var acValue = hadSnapshot
            ? snapshot.BatteryStandbyHibernateAcValue
            : TryGetCurrentHibernateAfterStandbyIndices(out var currentAcValue, out _)
                ? currentAcValue
                : HibernateAfterStandbyDisabledValue;
        var dcValue = hadSnapshot ? snapshot.BatteryStandbyHibernateDcValue : HibernateAfterStandbyFallbackRestoreDcValue;
        var failures = SetHibernateAfterStandbyPolicy(acValue, dcValue);

        if (failures.Length == 0)
        {
            snapshot.BatteryStandbyHibernateCaptured = false;
            snapshot.BatteryStandbyHibernateAcValue = 0;
            snapshot.BatteryStandbyHibernateDcValue = 0;
            RefreshBatteryStandbyHibernatePolicySummary();
            _logger.Info(_settings.BatteryStandbyHibernatePolicySummary);
            return;
        }

        RefreshBatteryStandbyHibernatePolicySummary();
        _logger.Warn($"恢复电池待机休眠策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshBatteryStandbyHibernatePolicySummary(bool persistSnapshot = true)
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var timeoutSeconds = GetBatteryStandbyHibernateTimeoutSeconds();
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            var status = EvaluateBatteryStandbyHibernateProtection();
            _settings.BatteryStandbyHibernatePolicySummary = status.Kind switch
            {
                ProtectionStatusKind.Applied => snapshot.BatteryStandbyHibernateCaptured
                    ? $"当前电源计划已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠（DC，可恢复到原始设置）"
                    : $"当前电源计划已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠（DC，历史基线缺失时将回退到永不休眠）",
                ProtectionStatusKind.Partial => $"已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠，但当前系统值更宽松",
                ProtectionStatusKind.Unknown => "已配置为电池待机后自动休眠，但暂时无法验证当前状态",
                _ => "已配置为电池待机后自动休眠，但当前未生效，请检查权限或日志"
            };
            if (persistSnapshot)
            {
                SaveSettingsSnapshot();
            }
            return;
        }

        if (TryGetCurrentHibernateAfterStandbyIndices(out var acValue, out var dcValue))
        {
            _settings.BatteryStandbyHibernatePolicySummary =
                $"未由应用接管；系统当前待机后休眠：AC={DescribePowerSettingDuration(acValue)}，DC={DescribePowerSettingDuration(dcValue)}";
        }
        else
        {
            _settings.BatteryStandbyHibernatePolicySummary = "未由应用接管；保留系统当前待机后休眠策略";
        }

        if (persistSnapshot)
        {
            SaveSettingsSnapshot();
        }
    }

    private void CaptureBatteryStandbyHibernateRestoreSnapshotIfNeeded()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        if (snapshot.BatteryStandbyHibernateCaptured)
        {
            return;
        }

        if (!TryGetCurrentHibernateAfterStandbyIndices(out var acValue, out var dcValue))
        {
            _logger.Warn("未能记录待机后休眠原始 AC/DC 设置；恢复时将回退到电池永不休眠。");
            return;
        }

        snapshot.BatteryStandbyHibernateAcValue = acValue;
        snapshot.BatteryStandbyHibernateDcValue = dcValue;
        snapshot.BatteryStandbyHibernateCaptured = true;
        SaveSettingsSnapshot();
        _logger.Info($"已记录待机后休眠原始设置：AC={acValue}，DC={dcValue}。");
    }

    private void ApplyKnownRemoteWakePolicy()
    {
        ApplyKnownRemoteWakePolicy(RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries));
    }

    private void ApplyKnownRemoteWakePolicy(IReadOnlyList<PowerRequestOverrideEntry> managedEntries)
    {
        var failures = new List<string>();

        foreach (var entry in managedEntries)
        {
            var result = ApplyRequestOverride(entry, FullRequestOverrideMask);
            if (!string.IsNullOrWhiteSpace(result))
            {
                failures.Add($"{entry.CallerTypeArgument}:{entry.Name}: {result}");
            }
        }

        var appliedCount = CountKnownRemoteRequestOverrides(managedEntries, static value => value == FullRequestOverrideMask);

        if (appliedCount == managedEntries.Count && failures.Count == 0)
        {
            _settings.KnownRemoteWakePolicySummary =
                _settings.KnownRemoteWakeRequestBackupCaptured
                    ? $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {appliedCount}/{managedEntries.Count} 条规则，可恢复到拦截前状态）"
                    : $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {appliedCount}/{managedEntries.Count} 条规则，历史基线缺失时将按清除处理）";
            SaveSettingsSnapshot();
            _logger.Info($"{_settings.KnownRemoteWakePolicySummary}。覆盖：{RemoteWakeBlockCatalog.ProductSummary(_settings.CustomRemoteWakeEntries)}。");
            return;
        }

        _settings.KnownRemoteWakePolicySummary = failures.Count == 0
            ? $"已尝试应用远程软件拦截规则，但当前仅命中 {appliedCount}/{managedEntries.Count} 条规则，请检查权限或日志"
            : $"应用远程软件拦截规则时收到输出，请检查日志（已命中 {appliedCount}/{managedEntries.Count} 条规则）";
        SaveSettingsSnapshot();
        if (failures.Count > 0)
        {
            _logger.Warn($"应用远程软件拦截规则时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
        }
    }

    private void RestoreKnownRemoteWakePolicy()
    {
        RestoreKnownRemoteWakePolicy(RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries));
    }

    private void RestoreKnownRemoteWakePolicy(IReadOnlyList<PowerRequestOverrideEntry> managedEntries)
    {
        var failures = new List<string>();
        var hadSnapshot = _settings.KnownRemoteWakeRequestBackupCaptured;
        failures.AddRange(RestoreKnownRemoteWakeEntries(managedEntries));
        var mismatchCount = CountKnownRemoteWakeRestoreMismatches(managedEntries, _settings.KnownRemoteWakeRequestOverrideBackup);

        if (failures.Count == 0 && mismatchCount == 0)
        {
            _settings.KnownRemoteWakePolicySummary = hadSnapshot
                ? "已恢复常见远程软件拦截前的系统请求替代状态"
                : "已清除应用管理的常见远程软件拦截规则（缺少历史基线时按清除处理）";
            _settings.KnownRemoteWakeRequestBackupCaptured = false;
            _settings.KnownRemoteWakeRequestOverrideBackup = [];
            SaveSettingsSnapshot();
            _logger.Info(_settings.KnownRemoteWakePolicySummary);
            return;
        }

        _settings.KnownRemoteWakePolicySummary = failures.Count == 0
            ? (hadSnapshot
                ? "已尝试恢复常见远程软件拦截前状态，但当前仍有规则未恢复到预期值"
                : "已尝试清除常见远程软件拦截规则，但当前仍检测到残留规则")
            : (hadSnapshot
                ? "尝试恢复常见远程软件拦截前状态时收到输出，请检查日志"
                : "尝试清除常见远程软件拦截规则时收到输出，请检查日志");
        SaveSettingsSnapshot();
        if (failures.Count > 0)
        {
            _logger.Warn($"恢复远程软件拦截规则时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
        }
    }

    private void RefreshKnownRemoteWakePolicySummary(bool persistSnapshot = true)
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);
        var managedCount = CountKnownRemoteRequestOverrides(managedEntries, static value => value > 0);
        var exactCount = CountKnownRemoteRequestOverrides(managedEntries, static value => value == FullRequestOverrideMask);

        if (_settings.BlockKnownRemoteWakeRequests)
        {
            _settings.KnownRemoteWakePolicySummary = exactCount == managedEntries.Count
                ? (_settings.KnownRemoteWakeRequestBackupCaptured
                    ? $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {exactCount}/{managedEntries.Count} 条规则，可恢复到拦截前状态）"
                    : $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {exactCount}/{managedEntries.Count} 条规则，历史基线缺失时将按清除处理）")
                : exactCount > 0
                    ? $"已配置为拦截常见远程软件的保持唤醒请求，但当前仅命中 {exactCount}/{managedEntries.Count} 条完整规则"
                    : managedCount > 0
                        ? $"已配置为拦截常见远程软件的保持唤醒请求，但当前仅检测到不完整的请求替代规则"
                        : "已配置为拦截常见远程软件的保持唤醒请求，但当前未检测到有效规则，请检查权限或日志";
        }
        else
        {
            _settings.KnownRemoteWakePolicySummary = managedCount == 0
                ? "未由应用接管；未检测到常见远程软件保持唤醒拦截规则"
                : $"未由应用接管；系统当前仍存在 {managedCount} 条常见远程软件请求替代";
        }

        if (persistSnapshot)
        {
            SaveSettingsSnapshot();
        }
    }

    private void CaptureKnownRemoteWakeRestoreSnapshotIfNeeded()
    {
        CaptureKnownRemoteWakeRestoreSnapshotIfNeeded(RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries));
    }

    private void CaptureKnownRemoteWakeRestoreSnapshotIfNeeded(IReadOnlyList<PowerRequestOverrideEntry> managedEntries)
    {
        var snapshot = new Dictionary<string, int>(
            _settings.KnownRemoteWakeRequestOverrideBackup,
            StringComparer.OrdinalIgnoreCase);
        var refreshedCount = 0;
        var removedStaleCount = 0;
        foreach (var entry in managedEntries)
        {
            var backupKey = BuildRequestOverrideBackupKey(entry);
            if (!TryGetRequestOverrideValue(entry, out var value))
            {
                if (snapshot.Remove(backupKey))
                {
                    removedStaleCount++;
                }

                continue;
            }

            snapshot[backupKey] = value;
            refreshedCount++;
        }

        _settings.KnownRemoteWakeRequestOverrideBackup = snapshot;
        _settings.KnownRemoteWakeRequestBackupCaptured = true;
        SaveSettingsSnapshot();
        _logger.Info(refreshedCount == 0 && removedStaleCount == 0
            ? $"已确认常见远程软件请求替代基线（现有规则 {snapshot.Count} 条）。"
            : $"已刷新常见远程软件请求替代基线 {refreshedCount} 条，清理过期基线 {removedStaleCount} 条（现有规则 {snapshot.Count} 条）。");
    }

    private void ReconcileKnownRemoteWakeManagedEntries(
        IReadOnlyList<PowerRequestOverrideEntry> previousManagedEntries,
        IReadOnlyList<PowerRequestOverrideEntry> currentManagedEntries)
    {
        var previousByKey = previousManagedEntries.ToDictionary(BuildRequestOverrideBackupKey, static entry => entry, StringComparer.OrdinalIgnoreCase);
        var currentByKey = currentManagedEntries.ToDictionary(BuildRequestOverrideBackupKey, static entry => entry, StringComparer.OrdinalIgnoreCase);

        var removedEntries = previousByKey
            .Where(pair => !currentByKey.ContainsKey(pair.Key))
            .Select(static pair => pair.Value)
            .ToArray();
        if (removedEntries.Length > 0)
        {
            var failures = RestoreKnownRemoteWakeEntries(removedEntries);
            if (failures.Count > 0)
            {
                _logger.Warn($"回滚已移除的远控拦截条目时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
            }
            else
            {
                foreach (var entry in removedEntries)
                {
                    _settings.KnownRemoteWakeRequestOverrideBackup.Remove(BuildRequestOverrideBackupKey(entry));
                }

                SaveSettingsSnapshot();
                _logger.Info($"已回滚 {removedEntries.Length} 条不再受管的远控拦截规则。");
            }
        }

        var addedEntries = currentByKey
            .Where(pair => !previousByKey.ContainsKey(pair.Key))
            .Select(static pair => pair.Value)
            .ToArray();
        if (addedEntries.Length > 0)
        {
            CaptureKnownRemoteWakeRestoreSnapshotIfNeeded(addedEntries);
        }
    }

    private List<string> RestoreKnownRemoteWakeEntries(IEnumerable<PowerRequestOverrideEntry> entries)
    {
        var failures = new List<string>();
        foreach (var entry in entries)
        {
            var backupKey = BuildRequestOverrideBackupKey(entry);
            if (_settings.KnownRemoteWakeRequestBackupCaptured
                && _settings.KnownRemoteWakeRequestOverrideBackup.TryGetValue(backupKey, out var requestMask)
                && requestMask > 0)
            {
                var result = ApplyRequestOverride(entry, requestMask);
                if (!string.IsNullOrWhiteSpace(result))
                {
                    failures.Add($"{entry.CallerTypeArgument}:{entry.Name}: {result}");
                }

                continue;
            }

            if (!TryRemoveRequestOverride(entry, out var failureMessage))
            {
                failures.Add(failureMessage!);
            }
        }

        return failures;
    }

    private int CountKnownRemoteWakeRestoreMismatches(
        IEnumerable<PowerRequestOverrideEntry> entries,
        IReadOnlyDictionary<string, int> backup)
    {
        var mismatchCount = 0;
        foreach (var entry in entries)
        {
            var backupKey = BuildRequestOverrideBackupKey(entry);
            if (backup.TryGetValue(backupKey, out var expectedMask) && expectedMask > 0)
            {
                if (!TryGetRequestOverrideValue(entry, out var restoredMask) || restoredMask != expectedMask)
                {
                    mismatchCount++;
                }

                continue;
            }

            if (TryGetRequestOverrideValue(entry, out var residualMask) && residualMask > 0)
            {
                mismatchCount++;
            }
        }

        return mismatchCount;
    }

    private void EnsureAutostartMatchesSettings()
    {
        var requireElevatedAutostart = RequiresElevatedAutostart();
        try
        {
            var status = AutostartManager.EnsureConfigured(_settings.StartWithWindows, requireElevatedAutostart);
            if (!string.Equals(_settings.AutostartPolicySummary, status.Summary, StringComparison.Ordinal))
            {
                _logger.Info(status.Summary);
            }

            _settings.AutostartPolicySummary = status.Summary;
            SaveSettingsSnapshot();
        }
        catch (Exception ex)
        {
            _settings.AutostartPolicySummary = $"开机自启配置失败：{ex.Message}";
            SaveSettingsSnapshot();
            _logger.Warn(_settings.AutostartPolicySummary);
        }
    }

    private void RefreshAutostartPolicySummary()
    {
        var requireElevatedAutostart = RequiresElevatedAutostart();
        var status = AutostartManager.QueryStatus(_settings.StartWithWindows, requireElevatedAutostart);
        if (!string.Equals(_settings.AutostartPolicySummary, status.Summary, StringComparison.Ordinal))
        {
            if (status.VerificationFailed)
            {
                _logger.Warn(status.Summary);
            }
            else
            {
                _logger.Info(status.Summary);
            }
        }

        _settings.AutostartPolicySummary = status.Summary;
    }

    private bool RequiresElevatedAutostart()
    {
        return _settings.DisableWakeTimers
            || _settings.DisableStandbyConnectivity
            || _settings.DisableWiFiDirectAdapters
            || _settings.EnforceBatteryStandbyHibernate
            || _settings.BlockKnownRemoteWakeRequests;
    }

    private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs e)
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            switch (e.Mode)
            {
                case PowerModes.Suspend:
                    CancelPendingResumeProtection();
                    CancelPendingResumeEvaluation();
                    ClearManualResumeSignals();
                    _settings.LastSuspendUtc = DateTimeOffset.UtcNow;
                    SaveSettingsSnapshot();
                    _logger.Warn("系统即将进入挂起/休眠。");
                    StateChanged?.Invoke(this, EventArgs.Empty);
                    break;

                case PowerModes.Resume:
                    var resumedAtUtc = DateTimeOffset.UtcNow;
                    var resumedAtTick = NativeMethods.GetTickCount();
                    _settings.LastResumeUtc = resumedAtUtc;
                    _settings.LastWakeSummary = "正在等待恢复证据…";
                    _settings.LastWakeEvidenceSummary = $"已收到恢复事件，等待约 {ResumeAnalysisDelayMilliseconds / 1000.0:F1} 秒后补抓更完整证据。";
                    SaveSettingsSnapshot();
                    _logger.Warn("系统已从挂起/休眠恢复，正在等待系统写入更完整的恢复证据。");
                    StateChanged?.Invoke(this, EventArgs.Empty);
                    QueueResumeEvaluation(resumedAtUtc, resumedAtTick);
                    break;

                case PowerModes.StatusChange:
                    if (!ReapplyManagedPoliciesForCurrentPowerPlan("电源状态变化"))
                    {
                        StateChanged?.Invoke(this, EventArgs.Empty);
                    }
                    break;
            }
        }
    }

    private void OnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
    {
        if (e.Category is not (UserPreferenceCategory.Power or UserPreferenceCategory.Policy))
        {
            return;
        }

        _ = ReapplyManagedPoliciesForCurrentPowerPlan($"系统偏好变化（{e.Category}）", forceReapplyForCurrentPlan: true);
    }

    private void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
    {
        lock (_stateSync)
        {
            if (e.Reason is not (SessionSwitchReason.SessionUnlock
                or SessionSwitchReason.ConsoleConnect
                or SessionSwitchReason.RemoteConnect
                or SessionSwitchReason.SessionLogon))
            {
                return;
            }

            InvalidateStatusSnapshot();
            var reason = DescribeSessionSwitchReason(e.Reason);
            RecordManualResumeSignal(reason);

            if (CancelPendingResumeProtection($"检测到{reason}，已取消待执行的自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。"))
            {
                StateChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    private void OnLidStateChanged(object? sender, bool isOpen)
    {
        lock (_stateSync)
        {
            if (!isOpen)
            {
                return;
            }

            InvalidateStatusSnapshot();
            RecordManualResumeSignal("开盖");

            if (CancelPendingResumeProtection($"检测到开盖操作，已取消待执行的自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。"))
            {
                StateChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    private void ArmResumeProtectionIfNeeded(WakeAnalysis analysis, uint resumedAtTick)
    {
        CancelPendingResumeProtection();

        if (_settings.PolicyMode != PowerPolicyMode.FollowPowerPlan)
        {
            return;
        }

        if (!_settings.ResumeProtectionEnabled)
        {
            return;
        }

        if (_settings.ResumeProtectionOnlyForUnattendedWake)
        {
            if (analysis.Kind == WakeKind.Manual)
            {
                _logger.Info($"本次恢复被判定为人工唤醒，已跳过自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
                return;
            }

            if (TryGetRecentManualResumeSignal(out var manualSignalReason))
            {
                _logger.Info($"检测到{manualSignalReason}，已跳过自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
                return;
            }
        }

        var delaySeconds = Math.Max(3, _settings.ResumeProtectionDelaySeconds);
        var elapsedMilliseconds = unchecked(NativeMethods.GetTickCount() - resumedAtTick);
        var requestedDelayMilliseconds = checked(delaySeconds * 1000);
        var remainingDelayMilliseconds = elapsedMilliseconds >= requestedDelayMilliseconds
            ? 0
            : requestedDelayMilliseconds - (int)elapsedMilliseconds;
        lock (_resumeProtectionSync)
        {
            _resumeProtectionArmedAtTick = NativeMethods.GetTickCount();
            _resumeTimer.Change(remainingDelayMilliseconds, Timeout.Infinite);
        }
        if (remainingDelayMilliseconds == 0)
        {
            _logger.Warn($"恢复证据补抓已完成，已立即触发自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
            return;
        }

        var remainingSeconds = Math.Max(1, (int)Math.Ceiling(remainingDelayMilliseconds / 1000d));
        _logger.Warn($"已启动恢复保护，将在恢复后的剩余 {remainingSeconds} 秒后自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
    }

    private void OnResumeTimerElapsed()
    {
        lock (_stateSync)
        {
            InvalidateStatusSnapshot();
            uint? armedAtTick;
            lock (_resumeProtectionSync)
            {
                _resumeTimer.Change(Timeout.Infinite, Timeout.Infinite);
                armedAtTick = _resumeProtectionArmedAtTick;
                _resumeProtectionArmedAtTick = null;
            }

            if (armedAtTick is null)
            {
                return;
            }

            if (HasUserInteractionSince(armedAtTick.Value))
            {
                _logger.Info($"检测到恢复后已有人工操作，已取消本次自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
                StateChanged?.Invoke(this, EventArgs.Empty);
                return;
            }

            if (_settings.ResumeProtectionMode == ResumeProtectionMode.Hibernate)
            {
                RequestSuspend(PowerState.Hibernate, "恢复保护到期，执行自动休眠。");
            }
            else if (_settings.ResumeProtectionMode == ResumeProtectionMode.LockScreen)
            {
                RequestLockWorkStation("恢复保护到期，执行自动锁屏。");
            }
            else
            {
                RequestSuspend(PowerState.Suspend, "恢复保护到期，执行自动睡眠。");
            }
        }
    }

    private static string DescribeMode(PowerPolicyMode mode)
    {
        return mode == PowerPolicyMode.KeepAwakeIndefinitely ? "无限期保持激活" : "遵循电源计划";
    }

    private static string DescribeResumeProtection(ResumeProtectionMode mode)
    {
        return mode switch
        {
            ResumeProtectionMode.Hibernate => "休眠",
            ResumeProtectionMode.LockScreen => "锁屏",
            _ => "睡眠"
        };
    }

    private string BuildCurrentStatus()
    {
        return BuildCurrentStatus(_settings);
    }

    private static string BuildCurrentStatus(AppSettings settings)
    {
        var delaySeconds = Math.Max(3, settings.ResumeProtectionDelaySeconds);
        return settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely
            ? "无限期保持激活"
            : (settings.ResumeProtectionEnabled
                ? $"遵循电源计划，恢复后 {delaySeconds} 秒自动{DescribeResumeProtection(settings.ResumeProtectionMode)}"
                : "遵循电源计划");
    }

    private string BuildProtectionRuleSummary()
    {
        return BuildProtectionRuleSummary(_settings);
    }

    private static string BuildProtectionRuleSummary(AppSettings settings)
    {
        if (settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            return $"当前处于无限保持唤醒模式，不执行恢复后自动{DescribeResumeProtection(settings.ResumeProtectionMode)}。";
        }

        if (!settings.ResumeProtectionEnabled)
        {
            return $"恢复保护已关闭，系统恢复后不会自动{DescribeResumeProtection(settings.ResumeProtectionMode)}。";
        }

        var resumeAction = DescribeResumeProtection(settings.ResumeProtectionMode);
        var delaySeconds = Math.Max(3, settings.ResumeProtectionDelaySeconds);

        if (settings.ResumeProtectionOnlyForUnattendedWake)
        {
            return $"人工行为（键盘、鼠标、开盖、解锁、控制台/远程接管、登录）恢复后跳过自动{resumeAction}；其他恢复（软件、定时器、设备、来源不明）会在 {delaySeconds} 秒后自动{resumeAction}。";
        }

        return $"不区分人工或非人工，系统每次恢复后都会在 {delaySeconds} 秒后自动{resumeAction}。";
    }

    private string RunPowerCfg(string arguments)
    {
        return _powerCfg.Run(arguments);
    }

    private string RunPowerShellScript(string script)
    {
        try
        {
            using var process = new Process();
            var encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
            process.StartInfo.FileName = WindowsPowerShellPath;
            process.StartInfo.Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedCommand}";
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            if (!process.WaitForExit(PowerShellTimeoutMilliseconds))
            {
                TryKillProcess(process);
                process.WaitForExit();
                var timeoutMessage = $"PowerShell 脚本超时（超过 {PowerShellTimeoutMilliseconds / 1000} 秒）";
                _logger.Warn(timeoutMessage);
                return timeoutMessage;
            }

            Task.WaitAll([outputTask, errorTask], 1000);
            var output = outputTask.IsCompletedSuccessfully ? outputTask.Result : string.Empty;
            var error = errorTask.IsCompletedSuccessfully ? errorTask.Result : string.Empty;

            if (process.ExitCode == 0)
            {
                return string.IsNullOrWhiteSpace(output) ? string.Empty : output.Trim();
            }

            var message = string.IsNullOrWhiteSpace(error) ? output.Trim() : error.Trim();
            return string.IsNullOrWhiteSpace(message)
                ? $"PowerShell 脚本失败，退出码 {process.ExitCode}"
                : message;
        }
        catch (Exception ex)
        {
            _logger.Error($"执行 PowerShell 脚本失败：{ex.Message}");
            return $"PowerShell 调用失败：{ex.Message}";
        }
    }

    private bool TryGetWiFiDirectAdapters(out List<WiFiDirectAdapterDevice> adapters, out string? failureMessage)
    {
        const string script = """
            $devices = @(Get-CimInstance Win32_PnPEntity |
                Where-Object { $_.PNPClass -eq 'Net' -and $_.Name -like 'Microsoft Wi-Fi Direct Virtual Adapter*' } |
                Sort-Object Name)
            foreach ($device in $devices) {
                $name = [string]$device.Name
                $instanceId = [string]$device.PNPDeviceID
                $errorCode = [int]$device.ConfigManagerErrorCode
                $disabled = if ($errorCode -eq 22) { 1 } else { 0 }
                "{0}`t{1}`t{2}`t{3}" -f $name, $instanceId, $disabled, $errorCode
            }
            """;

        var output = RunPowerShellScript(script);
        adapters = [];

        if (string.IsNullOrWhiteSpace(output))
        {
            failureMessage = null;
            return true;
        }

        foreach (var line in output.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var parts = line.Split('\t');
            if (parts.Length < 4)
            {
                failureMessage = output;
                adapters = [];
                return false;
            }

            if (!int.TryParse(parts[3], out var errorCode))
            {
                errorCode = 0;
            }

            adapters.Add(new WiFiDirectAdapterDevice
            {
                Name = parts[0],
                InstanceId = parts[1],
                IsDisabled = parts[2] == "1",
                ErrorCode = errorCode
            });
        }

        failureMessage = null;
        return true;
    }

    private string SetWiFiDirectAdapterEnabled(string instanceId, bool enable)
    {
        var escapedInstanceId = instanceId.Replace("'", "''", StringComparison.Ordinal);
        var command = enable ? "Enable-PnpDevice" : "Disable-PnpDevice";
        var script = $"""
            $ErrorActionPreference = 'Stop'
            {command} -InstanceId '{escapedInstanceId}' -Confirm:$false | Out-Null
            """;
        return RunPowerShellScript(script);
    }

    private static void TryKillProcess(Process process)
    {
        try
        {
            process.Kill(true);
        }
        catch (InvalidOperationException)
        {
        }
    }

    private static WakeAnalysis AnalyzeWake(WakeDiagnosticSnapshot snapshot)
    {
        var lastWake = snapshot.LastWakeText.ToLowerInvariant();
        var wakeTimers = snapshot.WakeTimersText.ToLowerInvariant();
        var powerRequests = snapshot.PowerRequestsText.ToLowerInvariant();
        var eventSummary = snapshot.EventSummary.ToLowerInvariant();
        var combined = string.Join(
            Environment.NewLine,
            snapshot.LastWakeText,
            snapshot.WakeTimersText,
            snapshot.PowerRequestsText,
            snapshot.EventSummary,
            snapshot.SleepStudySummary).ToLowerInvariant();

        if (ContainsAny(combined, ManualWakeIndicators)
            || eventSummary.Contains("input keyboard")
            || eventSummary.Contains("input mouse")
            || eventSummary.Contains("power button")
            || eventSummary.Contains("lid"))
        {
            return new WakeAnalysis(WakeKind.Manual, "疑似人工唤醒（按键/开盖/鼠标键盘）");
        }

        if (ContainsAny(combined, DeviceWakeIndicators))
        {
            return new WakeAnalysis(WakeKind.Unattended, "疑似设备或网络唤醒");
        }

        if (ContainsAny(combined, TimerWakeIndicators)
            || eventSummary.Contains("windowsupdateclient")
            || eventSummary.Contains("maintenance scheduler")
            || eventSummary.Contains("update orchestrator")
            || eventSummary.Contains("pdc task client")
            || (ContainsAny(wakeTimers, TimerWakeIndicators)
                && !ContainsAny(wakeTimers, NoActiveWakeTimerIndicators)))
        {
            return new WakeAnalysis(WakeKind.Unattended, "疑似软件或定时器唤醒");
        }

        if (ContainsAny(powerRequests, RemoteWakeIndicators))
        {
            return new WakeAnalysis(WakeKind.Unattended, "疑似远控软件或保活请求导致恢复");
        }

        if (ContainsAny(lastWake, NoWakeSourceIndicators)
            || ContainsAny(lastWake, NoWakeHistoryIndicators)
            || ContainsAny(wakeTimers, DiagnosticsUnavailableIndicators)
            || ContainsAny(wakeTimers, NoActiveWakeTimerIndicators))
        {
            return new WakeAnalysis(WakeKind.Unknown, "唤醒来源不明确");
        }

        return new WakeAnalysis(WakeKind.Unknown, "唤醒来源未知");
    }

    private static bool ContainsAny(string source, params string[] values)
    {
        return values.Any(source.Contains);
    }

    private AppSettings CloneCurrentSettings()
    {
        lock (_stateSync)
        {
            return _settings.Clone();
        }
    }

    private bool TryRunDeferredStartupRefresh(Action<PowerController> refreshAction)
    {
        lock (_stateSync)
        {
            if (_disposed)
            {
                return false;
            }

            refreshAction(this);
            return true;
        }
    }

    private void SaveSettingsSnapshot()
    {
        lock (_stateSync)
        {
            _settingsStore.Save(_settings.Clone());
        }
    }

    private string[] SetStandbyConnectivityPolicy(int connectivityAcValue, int connectivityDcValue, int disconnectedAcValue, int disconnectedDcValue)
    {
        var acConnectivityResult = RunPowerCfg($"/setacvalueindex scheme_current sub_none connectivityinstandby {connectivityAcValue}");
        var dcConnectivityResult = RunPowerCfg($"/setdcvalueindex scheme_current sub_none connectivityinstandby {connectivityDcValue}");
        var acDisconnectedResult = RunPowerCfg($"/setacvalueindex scheme_current sub_none disconnectedstandbymode {disconnectedAcValue}");
        var dcDisconnectedResult = RunPowerCfg($"/setdcvalueindex scheme_current sub_none disconnectedstandbymode {disconnectedDcValue}");
        var activateResult = RunPowerCfg("/setactive scheme_current");

        return new[] { acConnectivityResult, dcConnectivityResult, acDisconnectedResult, dcDisconnectedResult, activateResult }
            .Where(static x => !string.IsNullOrWhiteSpace(x))
            .ToArray();
    }

    private string[] SetWakeTimerPolicy(int acValue, int dcValue)
    {
        var acResult = RunPowerCfg($"/setacvalueindex scheme_current sub_sleep rtcwake {acValue}");
        var dcResult = RunPowerCfg($"/setdcvalueindex scheme_current sub_sleep rtcwake {dcValue}");
        var activateResult = RunPowerCfg("/setactive scheme_current");

        return new[] { acResult, dcResult, activateResult }
            .Where(static x => !string.IsNullOrWhiteSpace(x))
            .ToArray();
    }

    private string[] SetHibernateAfterStandbyPolicy(int acValue, int dcValue)
    {
        var acResult = RunPowerCfg($"/setacvalueindex scheme_current sub_sleep hibernateidle {acValue}");
        var dcResult = RunPowerCfg($"/setdcvalueindex scheme_current sub_sleep hibernateidle {dcValue}");
        var activateResult = RunPowerCfg("/setactive scheme_current");

        return new[] { acResult, dcResult, activateResult }
            .Where(static x => !string.IsNullOrWhiteSpace(x))
            .ToArray();
    }

    private bool TryGetCurrentWakeTimerIndices(out int acValue, out int dcValue)
    {
        acValue = WakeTimerEnabledValue;
        dcValue = WakeTimerEnabledValue;

        var output = RunPowerCfg("/q scheme_current sub_sleep rtcwake");
        if (string.IsNullOrWhiteSpace(output))
        {
            return false;
        }

        return TryParsePowerSettingIndex(output, CurrentAcPowerSettingRegex, WakeTimerEnabledValue, out acValue)
            && TryParsePowerSettingIndex(output, CurrentDcPowerSettingRegex, WakeTimerEnabledValue, out dcValue);
    }

    private bool TryGetCurrentHibernateAfterStandbyIndices(out int acValue, out int dcValue)
    {
        acValue = HibernateAfterStandbyDisabledValue;
        dcValue = HibernateAfterStandbyFallbackRestoreDcValue;

        var output = RunPowerCfg("/q scheme_current sub_sleep hibernateidle");
        if (string.IsNullOrWhiteSpace(output))
        {
            return false;
        }

        return TryParsePowerSettingIndex(output, CurrentAcPowerSettingRegex, HibernateAfterStandbyDisabledValue, out acValue)
            && TryParsePowerSettingIndex(output, CurrentDcPowerSettingRegex, HibernateAfterStandbyFallbackRestoreDcValue, out dcValue);
    }

    private bool TryGetCurrentStandbyConnectivityIndices(out int connectivityAcValue, out int connectivityDcValue, out int disconnectedAcValue, out int disconnectedDcValue)
    {
        connectivityAcValue = StandbyConnectivityManagedByWindowsValue;
        connectivityDcValue = StandbyConnectivityManagedByWindowsValue;
        disconnectedAcValue = DisconnectedStandbyModeNormalValue;
        disconnectedDcValue = DisconnectedStandbyModeNormalValue;

        var connectivityOutput = RunPowerCfg("/q scheme_current sub_none connectivityinstandby");
        var disconnectedOutput = RunPowerCfg("/q scheme_current sub_none disconnectedstandbymode");

        if (string.IsNullOrWhiteSpace(connectivityOutput) || string.IsNullOrWhiteSpace(disconnectedOutput))
        {
            return false;
        }

        return TryParsePowerSettingIndex(connectivityOutput, CurrentAcPowerSettingRegex, StandbyConnectivityManagedByWindowsValue, out connectivityAcValue)
            && TryParsePowerSettingIndex(connectivityOutput, CurrentDcPowerSettingRegex, StandbyConnectivityManagedByWindowsValue, out connectivityDcValue)
            && TryParsePowerSettingIndex(disconnectedOutput, CurrentAcPowerSettingRegex, DisconnectedStandbyModeNormalValue, out disconnectedAcValue)
            && TryParsePowerSettingIndex(disconnectedOutput, CurrentDcPowerSettingRegex, DisconnectedStandbyModeNormalValue, out disconnectedDcValue);
    }

    private static bool TryParsePowerSettingIndex(string output, Regex regex, int fallbackValue, out int value)
    {
        var match = regex.Match(output);
        if (match.Success
            && int.TryParse(match.Groups["value"].Value, System.Globalization.NumberStyles.HexNumber, null, out value))
        {
            return true;
        }

        value = fallbackValue;
        return false;
    }

    private int CountKnownRemoteRequestOverrides(IEnumerable<PowerRequestOverrideEntry> entries, Predicate<int> predicate)
    {
        var count = 0;
        foreach (var entry in entries)
        {
            if (!TryGetRequestOverrideValue(entry, out var value))
            {
                continue;
            }

            if (predicate(value))
            {
                count++;
            }
        }

        return count;
    }

    private string ApplyRequestOverride(PowerRequestOverrideEntry entry, int requestMask)
    {
        return RunPowerCfg(BuildRequestOverrideArguments(entry, requestMask));
    }

    private static string BuildRequestOverrideArguments(PowerRequestOverrideEntry entry, int requestMask)
    {
        return $"/requestsoverride {entry.CallerTypeArgument} \"{entry.Name}\" {BuildRequestOverrideRequestTypes(requestMask)}";
    }

    private static string BuildRequestOverrideRequestTypes(int requestMask)
    {
        var requestTypes = new List<string>(3);
        if ((requestMask & RequestDisplayMask) != 0)
        {
            requestTypes.Add("DISPLAY");
        }

        if ((requestMask & RequestSystemMask) != 0)
        {
            requestTypes.Add("SYSTEM");
        }

        if ((requestMask & RequestAwayModeMask) != 0)
        {
            requestTypes.Add("AWAYMODE");
        }

        return string.Join(' ', requestTypes);
    }

    private static string BuildRequestOverrideBackupKey(PowerRequestOverrideEntry entry)
    {
        return $"{entry.CallerTypeArgument}:{entry.Name}";
    }

    private void EnsureSettingsDefaults()
    {
        _settings.PowerPlanRestoreSnapshots = _settings.PowerPlanRestoreSnapshots is null
            ? new Dictionary<string, PowerPlanRestoreSnapshot>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, PowerPlanRestoreSnapshot>(_settings.PowerPlanRestoreSnapshots, StringComparer.OrdinalIgnoreCase);
        _settings.CustomRemoteWakeEntries = RemoteWakeBlockCatalog.NormalizeCustomEntries(_settings.CustomRemoteWakeEntries).ToList();
        _settings.WiFiDirectAdapterRestoreInstanceIds = (_settings.WiFiDirectAdapterRestoreInstanceIds ?? [])
            .Where(static instanceId => !string.IsNullOrWhiteSpace(instanceId))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        _settings.KnownRemoteWakeRequestOverrideBackup = _settings.KnownRemoteWakeRequestOverrideBackup is null
            ? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, int>(_settings.KnownRemoteWakeRequestOverrideBackup, StringComparer.OrdinalIgnoreCase);
        _settings.BatteryStandbyHibernateTimeoutSeconds = Math.Clamp(
            _settings.BatteryStandbyHibernateTimeoutSeconds <= 0
                ? DefaultBatteryStandbyHibernateTimeoutSeconds
                : _settings.BatteryStandbyHibernateTimeoutSeconds,
            180,
            14400);
        _settings.LastWakeSummary ??= "无";
        _settings.LastWakeEvidenceSummary ??= "无";
        _settings.WakeTimerPolicySummary ??= "未检查";
        _settings.StandbyConnectivityPolicySummary ??= "未检查";
        _settings.WiFiDirectAdapterPolicySummary ??= "未检查";
        _settings.BatteryStandbyHibernatePolicySummary ??= "未检查";
        _settings.KnownRemoteWakePolicySummary ??= "未检查";
        _settings.AutostartPolicySummary ??= "未检查";
    }

    private void MigrateLegacyRestoreSnapshotsIfNeeded()
    {
        var migrated = false;
        var snapshot = GetOrCreateActivePowerPlanSnapshot();

        if (_settings.WakeTimerRestoreSnapshotCaptured && !snapshot.WakeTimerCaptured)
        {
            snapshot.WakeTimerCaptured = true;
            snapshot.WakeTimerAcValue = _settings.WakeTimerRestoreAcValue;
            snapshot.WakeTimerDcValue = _settings.WakeTimerRestoreDcValue;
            migrated = true;
        }

        if (_settings.StandbyConnectivityRestoreSnapshotCaptured && !snapshot.StandbyConnectivityCaptured)
        {
            snapshot.StandbyConnectivityCaptured = true;
            snapshot.StandbyConnectivityAcValue = _settings.StandbyConnectivityRestoreAcValue;
            snapshot.StandbyConnectivityDcValue = _settings.StandbyConnectivityRestoreDcValue;
            snapshot.DisconnectedStandbyModeAcValue = _settings.DisconnectedStandbyModeRestoreAcValue;
            snapshot.DisconnectedStandbyModeDcValue = _settings.DisconnectedStandbyModeRestoreDcValue;
            migrated = true;
        }

        if (_settings.BatteryStandbyHibernateRestoreSnapshotCaptured && !snapshot.BatteryStandbyHibernateCaptured)
        {
            snapshot.BatteryStandbyHibernateCaptured = true;
            snapshot.BatteryStandbyHibernateAcValue = _settings.BatteryStandbyHibernateRestoreAcValue;
            snapshot.BatteryStandbyHibernateDcValue = _settings.BatteryStandbyHibernateRestoreDcValue;
            migrated = true;
        }

        if (migrated)
        {
            _settings.WakeTimerRestoreSnapshotCaptured = false;
            _settings.WakeTimerRestoreAcValue = 0;
            _settings.WakeTimerRestoreDcValue = 0;
            _settings.StandbyConnectivityRestoreSnapshotCaptured = false;
            _settings.StandbyConnectivityRestoreAcValue = 0;
            _settings.StandbyConnectivityRestoreDcValue = 0;
            _settings.DisconnectedStandbyModeRestoreAcValue = 0;
            _settings.DisconnectedStandbyModeRestoreDcValue = 0;
            _settings.BatteryStandbyHibernateRestoreSnapshotCaptured = false;
            _settings.BatteryStandbyHibernateRestoreAcValue = 0;
            _settings.BatteryStandbyHibernateRestoreDcValue = 0;
            _logger.Info($"已将旧版全局电源恢复快照迁移到当前电源计划：{snapshot.PlanName}。");
        }
    }

    private PowerPlanRestoreSnapshot GetOrCreateActivePowerPlanSnapshot()
    {
        var planId = GetActivePowerPlanKey(out var planName);
        if (!_settings.PowerPlanRestoreSnapshots.TryGetValue(planId, out var snapshot))
        {
            snapshot = new PowerPlanRestoreSnapshot
            {
                PlanName = planName
            };
            _settings.PowerPlanRestoreSnapshots[planId] = snapshot;
        }
        else if (string.IsNullOrWhiteSpace(snapshot.PlanName) && !string.IsNullOrWhiteSpace(planName))
        {
            snapshot.PlanName = planName;
        }

        return snapshot;
    }

    private string GetActivePowerPlanKey(out string planName)
    {
        if (TryGetActivePowerPlanInfo(out var planId, out planName))
        {
            return planId;
        }

        planName = "当前电源计划";
        return "scheme_current";
    }

    private bool TryGetActivePowerPlanInfo(out string planId, out string planName)
    {
        planId = string.Empty;
        planName = string.Empty;

        var output = RunPowerCfg("/getactivescheme");
        if (string.IsNullOrWhiteSpace(output))
        {
            return false;
        }

        var guidMatch = ActiveSchemeRegex.Match(output);
        if (!guidMatch.Success)
        {
            return false;
        }

        planId = guidMatch.Groups["guid"].Value;
        var nameMatch = ActiveSchemeNameRegex.Match(output);
        if (nameMatch.Success)
        {
            planName = nameMatch.Groups["name"].Value.Trim();
        }

        return true;
    }

    private void RememberActivePowerPlan()
    {
        var planKey = GetActivePowerPlanKey(out _);
        lock (_activePowerPlanSync)
        {
            _activePowerPlanKey = planKey;
        }
    }

    private bool ReapplyManagedPoliciesForCurrentPowerPlan(string reason, bool forceReapplyForCurrentPlan = false)
    {
        lock (_stateSync)
        {
            if (!TryGetActivePowerPlanInfo(out var planId, out var planName))
            {
                return false;
            }

            var planChanged = false;
            lock (_activePowerPlanSync)
            {
                planChanged = !string.Equals(_activePowerPlanKey, planId, StringComparison.OrdinalIgnoreCase);
                if (!planChanged && !forceReapplyForCurrentPlan)
                {
                    return false;
                }

                _activePowerPlanKey = planId;
            }

            InvalidateStatusSnapshot();
            _logger.Info(
                planChanged
                    ? $"检测到当前电源计划已切换为 {planName} ({planId})，正在重新应用当前计划相关策略。来源：{reason}。"
                    : $"检测到当前电源计划 {planName} ({planId}) 的配置可能已变化，正在重新核验并恢复当前计划相关策略。来源：{reason}。");

            if (_settings.DisableWakeTimers)
            {
                if (planChanged)
                {
                    CaptureWakeTimerRestoreSnapshotIfNeeded();
                }

                ApplyWakeTimerPolicy();
            }
            else
            {
                RefreshWakeTimerPolicySummary();
            }

            if (_settings.DisableStandbyConnectivity)
            {
                if (planChanged)
                {
                    CaptureStandbyConnectivityRestoreSnapshotIfNeeded();
                }

                ApplyStandbyConnectivityPolicy();
            }
            else
            {
                RefreshStandbyConnectivityPolicySummary();
            }

            if (_settings.EnforceBatteryStandbyHibernate)
            {
                if (planChanged)
                {
                    CaptureBatteryStandbyHibernateRestoreSnapshotIfNeeded();
                }

                ApplyBatteryStandbyHibernatePolicy();
            }
            else
            {
                RefreshBatteryStandbyHibernatePolicySummary();
            }

            StateChanged?.Invoke(this, EventArgs.Empty);
            return true;
        }
    }

    private static int NormalizeBatteryStandbyHibernateTimeoutSeconds(int timeoutSeconds)
    {
        return Math.Clamp(
            timeoutSeconds <= 0
                ? DefaultBatteryStandbyHibernateTimeoutSeconds
                : timeoutSeconds,
            180,
            14400);
    }

    private int GetBatteryStandbyHibernateTimeoutSeconds()
    {
        _settings.BatteryStandbyHibernateTimeoutSeconds = NormalizeBatteryStandbyHibernateTimeoutSeconds(
            _settings.BatteryStandbyHibernateTimeoutSeconds);
        return _settings.BatteryStandbyHibernateTimeoutSeconds;
    }

    private string BuildWakeDiagnosticsText(WakeDiagnosticSnapshot snapshot, bool includePowerRequests, bool includeSleepStudy)
    {
        var builder = new System.Text.StringBuilder();
        builder.AppendLine($"conclusion:{Environment.NewLine}{_wakeDiagnostics.SummarizeEvidence(snapshot)}");
        builder.AppendLine();
        builder.AppendLine($"lastwake:{Environment.NewLine}{snapshot.LastWakeText}");
        builder.AppendLine();
        builder.AppendLine($"waketimers:{Environment.NewLine}{snapshot.WakeTimersText}");
        builder.AppendLine();
        builder.AppendLine($"events:{Environment.NewLine}{snapshot.EventSummary}");

        if (includePowerRequests)
        {
            builder.AppendLine();
            builder.AppendLine($"requests:{Environment.NewLine}{snapshot.PowerRequestsText}");
            builder.AppendLine();
            builder.AppendLine($"requestsoverride:{Environment.NewLine}{snapshot.RequestOverridesText}");
        }

        if (includeSleepStudy)
        {
            builder.AppendLine();
            builder.AppendLine($"sleepstudy:{Environment.NewLine}{snapshot.SleepStudySummary}");
        }

        if (!string.IsNullOrWhiteSpace(snapshot.SuggestedRemoteWakeEntriesSummary))
        {
            builder.AppendLine();
            builder.AppendLine($"remote-suggestions:{Environment.NewLine}{snapshot.SuggestedRemoteWakeEntriesSummary}");
        }

        return builder.ToString().TrimEnd();
    }

    private string BuildCapabilitySummary(ProtectionStatus remoteWakeStatus, ProtectionStatus wifiDirectStatus)
    {
        return BuildCapabilitySummary(_settings, _settingsStore.LastSaveError, remoteWakeStatus, wifiDirectStatus);
    }

    private string BuildCapabilitySummary(
        AppSettings settings,
        string? lastSaveError,
        ProtectionStatus remoteWakeStatus,
        ProtectionStatus wifiDirectStatus)
    {
        var elevated = IsRunningElevated();
        var parts = new List<string>
        {
            elevated
                ? "当前以管理员权限运行"
                : "当前未以管理员权限运行"
        };

        if (!elevated)
        {
            parts.Add("远控拦截、Wi-Fi Direct 适配器接管、请求替代和部分 powercfg 写入可能受限");
        }

        if (settings.StartWithWindows && !string.IsNullOrWhiteSpace(settings.AutostartPolicySummary))
        {
            parts.Add($"开机自启：{settings.AutostartPolicySummary}");
        }

        if (settings.BlockKnownRemoteWakeRequests
            && remoteWakeStatus.Kind is ProtectionStatusKind.Partial or ProtectionStatusKind.Missing or ProtectionStatusKind.Unknown)
        {
            parts.Add("远控拦截最近一次应用未完全成功");
        }

        if (settings.DisableWiFiDirectAdapters
            && wifiDirectStatus.Kind is ProtectionStatusKind.Partial or ProtectionStatusKind.Missing or ProtectionStatusKind.Unknown)
        {
            parts.Add("Wi-Fi Direct 接管最近一次应用未完全成功");
        }

        if (!string.IsNullOrWhiteSpace(lastSaveError))
        {
            parts.Add("最近一次配置写入失败，当前变更可能仅在本次运行有效");
        }

        return string.Join("；", parts) + "。";
    }

    private StatusSnapshot BuildDeferredStatusSnapshot()
    {
        return new StatusSnapshot(
            DateTimeOffset.MinValue,
            BuildCurrentStatus(),
            BuildProtectionRuleSummary(),
            BuildDeferredCapabilitySummary(),
            BuildDeferredRiskSummary(),
            BuildDeferredPowerPlanSummary(),
            BuildManagedRemoteEntriesSummary(),
            BuildDeferredQuickState(_settings.DisableWakeTimers),
            BuildDeferredQuickState(_settings.DisableStandbyConnectivity),
            BuildDeferredQuickState(_settings.DisableWiFiDirectAdapters),
            BuildDeferredQuickState(_settings.EnforceBatteryStandbyHibernate),
            BuildDeferredRemoteWakeQuickState());
    }

    private StatusSnapshot BuildValidatedStatusSnapshot()
    {
        return BuildValidatedStatusSnapshot(_settings, _settingsStore.LastSaveError);
    }

    private StatusSnapshot BuildValidatedStatusSnapshot(AppSettings settings, string? lastSaveError)
    {
        var wakeTimerStatus = EvaluateWakeTimerProtection(settings);
        var standbyConnectivityStatus = EvaluateStandbyConnectivityProtection(settings);
        var wifiDirectStatus = EvaluateWiFiDirectProtection(settings);
        var batteryStandbyHibernateStatus = EvaluateBatteryStandbyHibernateProtection(settings);
        var remoteWakeStatus = EvaluateKnownRemoteWakeProtection(settings);

        return new StatusSnapshot(
            DateTimeOffset.UtcNow,
            BuildCurrentStatus(settings),
            BuildProtectionRuleSummary(settings),
            BuildCapabilitySummary(settings, lastSaveError, remoteWakeStatus, wifiDirectStatus),
            BuildRiskSummary(settings, lastSaveError, wakeTimerStatus, standbyConnectivityStatus, wifiDirectStatus, batteryStandbyHibernateStatus, remoteWakeStatus),
            BuildPowerPlanSummary(),
            BuildManagedRemoteEntriesSummary(settings),
            wakeTimerStatus.QuickState,
            standbyConnectivityStatus.QuickState,
            wifiDirectStatus.QuickState,
            batteryStandbyHibernateStatus.QuickState,
            remoteWakeStatus.QuickState);
    }

    private void QueueStatusSnapshotRefresh()
    {
        if (Interlocked.Exchange(ref _statusSnapshotRefreshQueued, 1) != 0)
        {
            return;
        }

        _ = Task.Run(() =>
        {
            var shouldNotify = false;
            var snapshotRevision = 0;

            try
            {
                AppSettings settingsSnapshot;
                string? lastSaveError;
                lock (_stateSync)
                {
                    if (_disposed || !_startupWarmupCompleted)
                    {
                        return;
                    }

                    settingsSnapshot = _settings.Clone();
                    lastSaveError = _settingsStore.LastSaveError;
                    snapshotRevision = _statusSnapshotRevision;
                }

                var snapshot = BuildValidatedStatusSnapshot(settingsSnapshot, lastSaveError);

                lock (_stateSync)
                {
                    if (_disposed || !_startupWarmupCompleted || snapshotRevision != _statusSnapshotRevision)
                    {
                        return;
                    }

                    lock (_statusSnapshotSync)
                    {
                        _statusSnapshot = snapshot;
                    }

                    shouldNotify = true;
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"后台状态校验失败：{ex.Message}");
            }
            finally
            {
                Interlocked.Exchange(ref _statusSnapshotRefreshQueued, 0);
            }

            if (shouldNotify)
            {
                StateChanged?.Invoke(this, EventArgs.Empty);
            }
        });
    }

    private string BuildDeferredCapabilitySummary()
    {
        var parts = new List<string>
        {
            IsRunningElevated()
                ? "当前以管理员权限运行"
                : "当前未以管理员权限运行"
        };

        if (_settings.StartWithWindows && !string.IsNullOrWhiteSpace(_settings.AutostartPolicySummary))
        {
            parts.Add($"开机自启：{_settings.AutostartPolicySummary}");
        }

        if (!string.IsNullOrWhiteSpace(_settingsStore.LastSaveError))
        {
            parts.Add("最近一次配置写入失败，当前变更可能仅在本次运行有效");
        }

        parts.Add("启动后详细校验尚未完成");
        return string.Join("；", parts) + "。";
    }

    private ProtectionStatus EvaluateWakeTimerProtection()
    {
        return EvaluateWakeTimerProtection(_settings);
    }

    private ProtectionStatus EvaluateWakeTimerProtection(AppSettings settings)
    {
        if (!TryGetCurrentWakeTimerIndices(out var acValue, out var dcValue))
        {
            return settings.DisableWakeTimers
                ? new ProtectionStatus(ProtectionStatusKind.Unknown, "当前：待验证")
                : new ProtectionStatus(ProtectionStatusKind.Missing, "当前：未接管");
        }

        var disabledCount = (acValue == WakeTimerDisabledValue ? 1 : 0) + (dcValue == WakeTimerDisabledValue ? 1 : 0);
        return disabledCount switch
        {
            2 => new ProtectionStatus(ProtectionStatusKind.Applied, settings.DisableWakeTimers ? "当前：已拦截" : "当前：系统已拦截"),
            > 0 => new ProtectionStatus(ProtectionStatusKind.Partial, settings.DisableWakeTimers ? "当前：部分拦截" : "当前：部分拦截"),
            _ => new ProtectionStatus(ProtectionStatusKind.Missing, settings.DisableWakeTimers ? "当前：未生效" : "当前：未接管")
        };
    }

    private ProtectionStatus EvaluateStandbyConnectivityProtection()
    {
        return EvaluateStandbyConnectivityProtection(_settings);
    }

    private ProtectionStatus EvaluateStandbyConnectivityProtection(AppSettings settings)
    {
        if (!TryGetCurrentStandbyConnectivityIndices(
                out var connectivityAcValue,
                out var connectivityDcValue,
                out var disconnectedAcValue,
                out var disconnectedDcValue))
        {
            return settings.DisableStandbyConnectivity
                ? new ProtectionStatus(ProtectionStatusKind.Unknown, "当前：待验证")
                : new ProtectionStatus(ProtectionStatusKind.Missing, "当前：未接管");
        }

        var matchCount =
            (connectivityAcValue == StandbyConnectivityDisabledValue ? 1 : 0) +
            (connectivityDcValue == StandbyConnectivityDisabledValue ? 1 : 0) +
            (disconnectedAcValue == DisconnectedStandbyModeAggressiveValue ? 1 : 0) +
            (disconnectedDcValue == DisconnectedStandbyModeAggressiveValue ? 1 : 0);
        return matchCount switch
        {
            4 => new ProtectionStatus(ProtectionStatusKind.Applied, settings.DisableStandbyConnectivity ? "当前：已拦截" : "当前：系统已拦截"),
            > 0 => new ProtectionStatus(ProtectionStatusKind.Partial, settings.DisableStandbyConnectivity ? "当前：部分拦截" : "当前：部分拦截"),
            _ => new ProtectionStatus(ProtectionStatusKind.Missing, settings.DisableStandbyConnectivity ? "当前：未生效" : "当前：未接管")
        };
    }

    private ProtectionStatus EvaluateWiFiDirectProtection()
    {
        return EvaluateWiFiDirectProtection(_settings);
    }

    private ProtectionStatus EvaluateWiFiDirectProtection(AppSettings settings)
    {
        if (!TryGetWiFiDirectAdapters(out var adapters, out _))
        {
            return settings.DisableWiFiDirectAdapters
                ? new ProtectionStatus(ProtectionStatusKind.Unknown, "当前：待验证")
                : new ProtectionStatus(ProtectionStatusKind.Missing, "当前：未接管");
        }

        if (adapters.Count == 0)
        {
            return new ProtectionStatus(ProtectionStatusKind.NotApplicable, "当前：无设备");
        }

        var disabledCount = adapters.Count(static adapter => adapter.IsDisabled);
        if (disabledCount == adapters.Count)
        {
            return new ProtectionStatus(ProtectionStatusKind.Applied, settings.DisableWiFiDirectAdapters ? "当前：已禁用" : "当前：系统已禁用");
        }

        if (disabledCount > 0)
        {
            return new ProtectionStatus(ProtectionStatusKind.Partial, settings.DisableWiFiDirectAdapters ? "当前：部分禁用" : "当前：部分禁用");
        }

        return new ProtectionStatus(ProtectionStatusKind.Missing, settings.DisableWiFiDirectAdapters ? "当前：未生效" : "当前：未接管");
    }

    private ProtectionStatus EvaluateBatteryStandbyHibernateProtection()
    {
        return EvaluateBatteryStandbyHibernateProtection(_settings);
    }

    private ProtectionStatus EvaluateBatteryStandbyHibernateProtection(AppSettings settings)
    {
        if (!TryGetCurrentHibernateAfterStandbyIndices(out _, out var dcValue))
        {
            return settings.EnforceBatteryStandbyHibernate
                ? new ProtectionStatus(ProtectionStatusKind.Unknown, "当前：待验证")
                : new ProtectionStatus(ProtectionStatusKind.Missing, "当前：未接管");
        }

        var timeoutSeconds = NormalizeBatteryStandbyHibernateTimeoutSeconds(settings.BatteryStandbyHibernateTimeoutSeconds);
        if (dcValue > 0 && dcValue != HibernateAfterStandbyFallbackRestoreDcValue && dcValue <= timeoutSeconds)
        {
            return new ProtectionStatus(ProtectionStatusKind.Applied, settings.EnforceBatteryStandbyHibernate ? "当前：已兜底" : "当前：系统已兜底");
        }

        if (dcValue > 0 && dcValue != HibernateAfterStandbyFallbackRestoreDcValue)
        {
            return new ProtectionStatus(ProtectionStatusKind.Partial, settings.EnforceBatteryStandbyHibernate ? "当前：兜底较弱" : "当前：系统已兜底");
        }

        return new ProtectionStatus(ProtectionStatusKind.Missing, settings.EnforceBatteryStandbyHibernate ? "当前：未生效" : "当前：未接管");
    }

    private ProtectionStatus EvaluateKnownRemoteWakeProtection()
    {
        return EvaluateKnownRemoteWakeProtection(_settings);
    }

    private ProtectionStatus EvaluateKnownRemoteWakeProtection(AppSettings settings)
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(settings.CustomRemoteWakeEntries);
        if (managedEntries.Count == 0)
        {
            return new ProtectionStatus(ProtectionStatusKind.NotApplicable, "当前：无规则");
        }

        var appliedCount = CountKnownRemoteRequestOverrides(managedEntries, static value => value == FullRequestOverrideMask);
        if (appliedCount == managedEntries.Count)
        {
            return new ProtectionStatus(ProtectionStatusKind.Applied, settings.BlockKnownRemoteWakeRequests ? "当前：已拦截" : "当前：系统已拦截");
        }

        if (appliedCount > 0)
        {
            return new ProtectionStatus(ProtectionStatusKind.Partial, settings.BlockKnownRemoteWakeRequests ? "当前：部分拦截" : "当前：部分拦截");
        }

        return new ProtectionStatus(ProtectionStatusKind.Missing, settings.BlockKnownRemoteWakeRequests ? "当前：未生效" : "当前：未接管");
    }

    private string BuildRiskSummary(
        AppSettings settings,
        string? lastSaveError,
        ProtectionStatus wakeTimerStatus,
        ProtectionStatus standbyConnectivityStatus,
        ProtectionStatus wifiDirectStatus,
        ProtectionStatus batteryStandbyHibernateStatus,
        ProtectionStatus remoteWakeStatus)
    {
        var risks = new List<string>();

        if (settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            risks.Add("当前处于无限保持唤醒模式");
        }
        else
        {
            if (!settings.ResumeProtectionEnabled)
            {
                risks.Add("恢复保护已关闭");
            }

            if (settings.DisableWakeTimers
                && wakeTimerStatus.Kind is ProtectionStatusKind.Missing or ProtectionStatusKind.Partial or ProtectionStatusKind.Unknown)
            {
                risks.Add("唤醒定时器拦截未完全生效");
            }

            if (settings.DisableStandbyConnectivity
                && standbyConnectivityStatus.Kind is ProtectionStatusKind.Missing or ProtectionStatusKind.Partial or ProtectionStatusKind.Unknown)
            {
                risks.Add("待机联网拦截未完全生效");
            }

            if (settings.DisableWiFiDirectAdapters
                && wifiDirectStatus.Kind is ProtectionStatusKind.Missing or ProtectionStatusKind.Partial or ProtectionStatusKind.Unknown)
            {
                risks.Add("Wi-Fi Direct 稳态未完全生效");
            }

            if (settings.EnforceBatteryStandbyHibernate
                && batteryStandbyHibernateStatus.Kind is ProtectionStatusKind.Missing or ProtectionStatusKind.Partial or ProtectionStatusKind.Unknown)
            {
                risks.Add("电池兜底休眠未完全生效");
            }
        }

        if (settings.BlockKnownRemoteWakeRequests
            && remoteWakeStatus.Kind is ProtectionStatusKind.Missing or ProtectionStatusKind.Partial or ProtectionStatusKind.Unknown)
        {
            risks.Add("远控保活拦截未完全生效");
        }

        if (!string.IsNullOrWhiteSpace(lastSaveError))
        {
            risks.Add("最近一次配置写入失败，重启后部分变更可能丢失");
        }

        return risks.Count == 0
            ? "当前保护层完整，没有明显高风险缺口。"
            : $"当前最主要风险：{string.Join("；", risks.Take(3))}。";
    }

    private string BuildDeferredRiskSummary()
    {
        var risks = new List<string>();

        if (_settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            risks.Add("当前处于无限保持唤醒模式");
        }
        else if (!_settings.ResumeProtectionEnabled)
        {
            risks.Add("恢复保护已关闭");
        }

        if (!string.IsNullOrWhiteSpace(_settingsStore.LastSaveError))
        {
            risks.Add("最近一次配置写入失败，重启后部分变更可能丢失");
        }

        risks.Add("启动后详细校验尚未完成");

        return $"当前最主要风险：{string.Join("；", risks.Take(3))}。";
    }

    private string BuildDeferredPowerPlanSummary()
    {
        return _startupWarmupCompleted
            ? "当前电源计划：正在后台校验"
            : "当前电源计划：启动后后台校验中";
    }

    private static string BuildDeferredQuickState(bool enabled)
    {
        return enabled ? "当前：待校验" : "当前：未接管";
    }

    private string BuildDeferredRemoteWakeQuickState()
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);
        if (managedEntries.Count == 0)
        {
            return "当前：无规则";
        }

        return _settings.BlockKnownRemoteWakeRequests ? "当前：待校验" : "当前：未接管";
    }

    private string BuildPowerPlanSummary()
    {
        return TryGetActivePowerPlanInfo(out var planId, out var planName)
            ? $"当前电源计划：{planName} ({planId})"
            : "当前电源计划：未能识别 GUID";
    }

    private string BuildManagedRemoteEntriesSummary()
    {
        return BuildManagedRemoteEntriesSummary(_settings);
    }

    private static string BuildManagedRemoteEntriesSummary(AppSettings settings)
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(settings.CustomRemoteWakeEntries);
        var customCount = managedEntries.Count(static entry => entry.Product == "自定义");
        var builtInCount = managedEntries.Count - customCount;
        return customCount == 0
            ? $"内置远控拦截名单 {builtInCount} 条"
            : $"内置 {builtInCount} 条，自定义 {customCount} 条，共 {managedEntries.Count} 条";
    }

    private static string DescribeStandbyConnectivityValue(int value)
    {
        return value switch
        {
            StandbyConnectivityDisabledValue => "禁用",
            StandbyConnectivityEnabledValue => "启用",
            StandbyConnectivityManagedByWindowsValue => "由 Windows 管理",
            _ => $"未知({value})"
        };
    }

    private static string DescribeDisconnectedStandbyModeValue(int value)
    {
        return value switch
        {
            DisconnectedStandbyModeNormalValue => "正常",
            DisconnectedStandbyModeAggressiveValue => "主动",
            _ => $"未知({value})"
        };
    }

    private static string DescribePowerSettingDuration(int value)
    {
        if (value <= 0 || value == int.MaxValue)
        {
            return "永不";
        }

        if (value % 3600 == 0)
        {
            return $"{value / 3600} 小时";
        }

        if (value % 60 == 0)
        {
            return $"{value / 60} 分钟";
        }

        return $"{value} 秒";
    }

    private bool TryGetRequestOverrideValue(PowerRequestOverrideEntry entry, out int value)
    {
        value = 0;

        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(GetRequestOverrideRegistryPath(entry.CallerType), writable: false);
            if (key?.GetValue(entry.Name) is not { } rawValue)
            {
                return false;
            }

            value = Convert.ToInt32(rawValue);
            return true;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException or InvalidCastException or FormatException or OverflowException or System.Security.SecurityException)
        {
            _logger.Warn($"读取请求替代值失败：{entry.CallerTypeArgument}:{entry.Name}: {ex.Message}");
            return false;
        }
    }

    private static string GetRequestOverrideRegistryPath(PowerRequestOverrideCallerType callerType)
    {
        return Path.Combine(PowerRequestOverrideRegistryRoot, callerType.ToString());
    }

    private bool TryRemoveRequestOverride(PowerRequestOverrideEntry entry, out string? failureMessage)
    {
        failureMessage = null;

        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(GetRequestOverrideRegistryPath(entry.CallerType), writable: true);

            if (key is null)
            {
                return true;
            }

            key.DeleteValue(entry.Name, false);
            return true;
        }
        catch (UnauthorizedAccessException ex)
        {
            failureMessage = $"{entry.CallerTypeArgument}:{entry.Name}: {ex.Message}";
            return false;
        }
        catch (System.Security.SecurityException ex)
        {
            failureMessage = $"{entry.CallerTypeArgument}:{entry.Name}: {ex.Message}";
            return false;
        }
        catch (IOException ex)
        {
            failureMessage = $"{entry.CallerTypeArgument}:{entry.Name}: {ex.Message}";
            return false;
        }
    }

    private void RecordManualResumeSignal(string reason)
    {
        lock (_manualSignalSync)
        {
            _lastManualResumeSignalUtc = DateTimeOffset.UtcNow;
            _lastManualResumeSignalReason = reason;
        }
    }

    private void CancelPendingResumeEvaluation()
    {
        lock (_resumeEvaluationSync)
        {
            _resumeEvaluationGeneration++;
        }
    }

    private void QueueResumeEvaluation(DateTimeOffset resumedAtUtc, uint resumedAtTick)
    {
        int evaluationGeneration;
        lock (_resumeEvaluationSync)
        {
            _resumeEvaluationGeneration++;
            evaluationGeneration = _resumeEvaluationGeneration;
        }

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(ResumeAnalysisDelayMilliseconds).ConfigureAwait(false);
                lock (_resumeEvaluationSync)
                {
                    if (_disposed || evaluationGeneration != _resumeEvaluationGeneration)
                    {
                        return;
                    }
                }

                var customRemoteEntries = CloneCurrentSettings().CustomRemoteWakeEntries;
                var snapshot = _wakeDiagnostics.CollectSnapshot(
                    resumedAtUtc.AddMinutes(-10),
                    includePowerRequests: true,
                    includeSleepStudy: false,
                    existingCustomRemoteEntries: customRemoteEntries);
                lock (_resumeEvaluationSync)
                {
                    if (_disposed || evaluationGeneration != _resumeEvaluationGeneration)
                    {
                        return;
                    }
                }

                var analysis = AnalyzeWake(snapshot);

                if (TryGetRecentManualResumeSignal(out var manualSignalReason) && analysis.Kind != WakeKind.Manual)
                {
                    analysis = new WakeAnalysis(WakeKind.Manual, $"检测到人工行为（{manualSignalReason}）");
                }
                else if (HasUserInteractionSince(resumedAtTick) && analysis.Kind == WakeKind.Unknown)
                {
                    analysis = new WakeAnalysis(WakeKind.Manual, "检测到恢复后立即有人机输入");
                }

                lock (_stateSync)
                {
                    InvalidateStatusSnapshot();
                    _settings.LastWakeSummary = analysis.Summary;
                    _settings.LastWakeEvidenceSummary = _wakeDiagnostics.SummarizeEvidence(snapshot);
                    SaveSettingsSnapshot();
                    _logger.Warn($"系统已完成恢复分析。判定：{analysis.Summary}{Environment.NewLine}{BuildWakeDiagnosticsText(snapshot, includePowerRequests: true, includeSleepStudy: false)}");
                    StateChanged?.Invoke(this, EventArgs.Empty);
                    ArmResumeProtectionIfNeeded(analysis, resumedAtTick);
                }
            }
            catch (Exception ex)
            {
                lock (_stateSync)
                {
                    InvalidateStatusSnapshot();
                    _logger.Error($"恢复后分析失败：{ex.Message}");
                    _settings.LastWakeSummary = "恢复分析失败";
                    _settings.LastWakeEvidenceSummary = ex.Message;
                    SaveSettingsSnapshot();
                    StateChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        });
    }

    private void ClearManualResumeSignals()
    {
        lock (_manualSignalSync)
        {
            _lastManualResumeSignalUtc = null;
            _lastManualResumeSignalReason = "人工操作";
        }
    }

    private bool TryGetRecentManualResumeSignal(out string reason)
    {
        lock (_manualSignalSync)
        {
            if (_lastManualResumeSignalUtc is not { } lastSignalUtc)
            {
                reason = string.Empty;
                return false;
            }

            if (DateTimeOffset.UtcNow - lastSignalUtc > ManualResumeSignalWindow)
            {
                _lastManualResumeSignalUtc = null;
                _lastManualResumeSignalReason = "人工操作";
                reason = string.Empty;
                return false;
            }

            reason = _lastManualResumeSignalReason;
            return true;
        }
    }

    private bool CancelPendingResumeProtection(string? reason = null)
    {
        var canceled = false;

        lock (_resumeProtectionSync)
        {
            _resumeTimer.Change(Timeout.Infinite, Timeout.Infinite);

            if (_resumeProtectionArmedAtTick is not null)
            {
                _resumeProtectionArmedAtTick = null;
                canceled = true;
            }
        }

        if (canceled && !string.IsNullOrWhiteSpace(reason))
        {
            _logger.Info(reason);
        }

        return canceled;
    }

    private static bool HasUserInteractionSince(uint armedAtTick)
    {
        var lastInputInfo = new NativeMethods.LastInputInfo
        {
            cbSize = (uint)Marshal.SizeOf<NativeMethods.LastInputInfo>()
        };

        if (!NativeMethods.GetLastInputInfo(ref lastInputInfo))
        {
            return false;
        }

        var elapsed = unchecked(lastInputInfo.dwTime - armedAtTick);
        return elapsed != 0 && elapsed < 0x80000000u;
    }

    private void RequestSuspend(PowerState powerState, string logMessage)
    {
        InvalidateStatusSnapshot();
        _logger.Warn(logMessage);
        if (Application.SetSuspendState(powerState, false, false))
        {
            return;
        }

        _logger.Error($"请求{DescribePowerState(powerState)}失败，系统拒绝执行。");
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    private void RequestLockWorkStation(string logMessage)
    {
        InvalidateStatusSnapshot();
        _logger.Warn(logMessage);
        if (NativeMethods.LockWorkStation())
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        _logger.Error($"请求锁屏失败，系统拒绝执行。Win32Error={Marshal.GetLastWin32Error()}");
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    private static string DescribePowerState(PowerState powerState)
    {
        return powerState == PowerState.Hibernate ? "休眠" : "睡眠";
    }

    private static string DescribeSessionSwitchReason(SessionSwitchReason reason)
    {
        return reason switch
        {
            SessionSwitchReason.SessionUnlock => "会话解锁",
            SessionSwitchReason.ConsoleConnect => "控制台接管",
            SessionSwitchReason.RemoteConnect => "远程接管",
            SessionSwitchReason.SessionLogon => "用户登录",
            _ => reason.ToString()
        };
    }

    private static bool IsRunningElevated()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static string ResolveWindowsPowerShellPath()
    {
        var candidate = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe");
        return File.Exists(candidate) ? candidate : "powershell";
    }

    private enum ProtectionStatusKind
    {
        Applied,
        Partial,
        Missing,
        NotApplicable,
        Unknown
    }

    private readonly record struct StatusSnapshot(
        DateTimeOffset GeneratedAtUtc,
        string CurrentStatus,
        string ProtectionRuleSummary,
        string CapabilitySummary,
        string RiskSummary,
        string PowerPlanSummary,
        string ManagedRemoteEntriesSummary,
        string WakeTimerQuickState,
        string StandbyConnectivityQuickState,
        string WiFiDirectQuickState,
        string BatteryStandbyHibernateQuickState,
        string RemoteWakeQuickState)
    {
        public static StatusSnapshot Empty { get; } = new(
            DateTimeOffset.MinValue,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty);
    }

    private readonly record struct ProtectionStatus(ProtectionStatusKind Kind, string QuickState);

    private enum WakeKind
    {
        Manual,
        Unattended,
        Unknown
    }

    private readonly record struct WakeAnalysis(WakeKind Kind, string Summary);
}
