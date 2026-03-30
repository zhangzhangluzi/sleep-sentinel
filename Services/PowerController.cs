using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class PowerController : IDisposable
{
    private static readonly TimeSpan ManualResumeSignalWindow = TimeSpan.FromSeconds(20);
    private const int DefaultBatteryStandbyHibernateTimeoutSeconds = 600;
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
    private readonly object _manualSignalSync = new();
    private readonly object _resumeProtectionSync = new();
    private readonly PowerNotificationWindow? _powerNotificationWindow;
    private readonly System.Threading.Timer _resumeTimer;
    private DateTimeOffset? _lastManualResumeSignalUtc;
    private string _lastManualResumeSignalReason = "人工操作";
    private uint? _resumeProtectionArmedAtTick;
    private AppSettings _settings;

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
        ApplyPolicy(_settings.PolicyMode);
        if (_settings.DisableWakeTimers)
        {
            ApplyWakeTimerPolicy();
        }
        else
        {
            RefreshWakeTimerPolicySummary();
        }
        if (_settings.DisableStandbyConnectivity)
        {
            ApplyStandbyConnectivityPolicy();
        }
        else
        {
            RefreshStandbyConnectivityPolicySummary();
        }
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            CaptureBatteryStandbyHibernateRestoreSnapshotIfNeeded();
            ApplyBatteryStandbyHibernatePolicy();
        }
        else
        {
            RefreshBatteryStandbyHibernatePolicySummary();
        }
        if (_settings.BlockKnownRemoteWakeRequests)
        {
            ApplyKnownRemoteWakePolicy();
        }
        else
        {
            RefreshKnownRemoteWakePolicySummary();
        }
        EnsureAutostartMatchesSettings();
        _settingsStore.Save(_settings);
        _logger.Info($"应用启动，当前模式：{DescribeMode(_settings.PolicyMode)}。");
    }

    public event EventHandler? StateChanged;

    public AppSettings CurrentSettings => _settings;

    public string CurrentStatus =>
        _settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely
            ? "无限期保持激活"
            : (_settings.ResumeProtectionEnabled
                ? $"遵循电源计划，恢复后 {_settings.ResumeProtectionDelaySeconds} 秒自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}"
                : "遵循电源计划");

    public string CurrentProtectionRuleSummary => BuildProtectionRuleSummary();

    public string CurrentCapabilitySummary => BuildCapabilitySummary();

    public string CurrentRiskSummary => BuildRiskSummary();

    public string CurrentPowerPlanSummary => BuildPowerPlanSummary();

    public string CurrentManagedRemoteEntriesSummary => BuildManagedRemoteEntriesSummary();

    public void SetPolicyMode(PowerPolicyMode mode)
    {
        if (_settings.PolicyMode == mode)
        {
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.PolicyMode = mode;
        UpdateSettings(updatedSettings);
    }

    public void UpdateSettings(AppSettings updatedSettings)
    {
        CancelPendingResumeProtection();
        var previousSettings = _settings;
        _settings = updatedSettings;
        EnsureSettingsDefaults();
        _settingsStore.Save(_settings);
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
            RefreshWakeTimerPolicySummary();
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
            RefreshStandbyConnectivityPolicySummary();
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
            RefreshBatteryStandbyHibernatePolicySummary();
        }
        if (_settings.BlockKnownRemoteWakeRequests)
        {
            if (!previousSettings.BlockKnownRemoteWakeRequests)
            {
                CaptureKnownRemoteWakeRestoreSnapshotIfNeeded();
            }

            ApplyKnownRemoteWakePolicy();
        }
        else if (previousSettings.BlockKnownRemoteWakeRequests)
        {
            RestoreKnownRemoteWakePolicy();
        }
        else
        {
            RefreshKnownRemoteWakePolicySummary();
        }
        EnsureAutostartMatchesSettings();
        _logger.Info($"设置已更新：模式={DescribeMode(_settings.PolicyMode)}，恢复保护={_settings.ResumeProtectionEnabled}，待机联网拦截={_settings.DisableStandbyConnectivity}，电池兜底休眠={_settings.EnforceBatteryStandbyHibernate}（{DescribePowerSettingDuration(GetBatteryStandbyHibernateTimeoutSeconds())}），远控拦截={_settings.BlockKnownRemoteWakeRequests}，自定义远控={_settings.CustomRemoteWakeEntries.Count} 条。");
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SleepNow()
    {
        CancelPendingResumeProtection("用户手动操作，已取消待执行的自动回睡。");
        RequestSuspend(PowerState.Suspend, "用户手动请求立即睡眠。");
    }

    public void HibernateNow()
    {
        CancelPendingResumeProtection("用户手动操作，已取消待执行的自动回睡。");
        RequestSuspend(PowerState.Hibernate, "用户手动请求立即休眠。");
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
        return _wakeDiagnostics.CollectSnapshot(_settings.LastSuspendUtc?.AddMinutes(-10), includePowerRequests, includeSleepStudy);
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
        return _wakeDiagnostics.SuggestCustomRemoteWakeEntries(_settings.CustomRemoteWakeEntries);
    }

    public void UpdateCustomRemoteWakeEntries(IEnumerable<string> rawEntries)
    {
        var normalized = RemoteWakeBlockCatalog.NormalizeCustomEntries(rawEntries);
        if (_settings.CustomRemoteWakeEntries.SequenceEqual(normalized, StringComparer.OrdinalIgnoreCase))
        {
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.CustomRemoteWakeEntries = normalized.ToList();
        _logger.Info(normalized.Count == 0
            ? "已清空自定义远控拦截名单。"
            : $"已更新自定义远控拦截名单，共 {normalized.Count} 条。");
        UpdateSettings(updatedSettings);
    }

    public void ReapplyAllManagedSettings()
    {
        CancelPendingResumeProtection();
        ApplyPolicy(_settings.PolicyMode);

        if (_settings.DisableWakeTimers)
        {
            ApplyWakeTimerPolicy();
        }
        else
        {
            RefreshWakeTimerPolicySummary();
        }

        if (_settings.DisableStandbyConnectivity)
        {
            ApplyStandbyConnectivityPolicy();
        }
        else
        {
            RefreshStandbyConnectivityPolicySummary();
        }

        if (_settings.EnforceBatteryStandbyHibernate)
        {
            ApplyBatteryStandbyHibernatePolicy();
        }
        else
        {
            RefreshBatteryStandbyHibernatePolicySummary();
        }

        if (_settings.BlockKnownRemoteWakeRequests)
        {
            ApplyKnownRemoteWakePolicy();
        }
        else
        {
            RefreshKnownRemoteWakePolicySummary();
        }

        EnsureAutostartMatchesSettings();
        _settingsStore.Save(_settings);
        _logger.Info("已重新应用当前全部设置。");
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void ReapplyWakeTimerPolicy()
    {
        if (_settings.DisableWakeTimers)
        {
            ApplyWakeTimerPolicy();
        }
        else
        {
            _settings.WakeTimerPolicySummary = "未由应用接管；勾选后才会禁用唤醒定时器";
            _settingsStore.Save(_settings);
            _logger.Info(_settings.WakeTimerPolicySummary);
        }

        _settingsStore.Save(_settings);
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void ReapplyKnownRemoteWakePolicy()
    {
        if (_settings.BlockKnownRemoteWakeRequests)
        {
            ApplyKnownRemoteWakePolicy();
        }
        else
        {
            RefreshKnownRemoteWakePolicySummary();
            _settingsStore.Save(_settings);
        }

        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void ReapplyStandbyConnectivityPolicy()
    {
        if (_settings.DisableStandbyConnectivity)
        {
            ApplyStandbyConnectivityPolicy();
        }
        else
        {
            RefreshStandbyConnectivityPolicySummary();
            _settingsStore.Save(_settings);
            _logger.Info(_settings.StandbyConnectivityPolicySummary);
        }

        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void ReapplyBatteryStandbyHibernatePolicy()
    {
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            ApplyBatteryStandbyHibernatePolicy();
        }
        else
        {
            RefreshBatteryStandbyHibernatePolicySummary();
            _settingsStore.Save(_settings);
            _logger.Info(_settings.BatteryStandbyHibernatePolicySummary);
        }

        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void BlockSoftwareWake()
    {
        if (_settings.DisableWakeTimers)
        {
            _logger.Info("软件/定时器唤醒拦截已启用，正在重新应用当前电源计划策略。");
            ApplyWakeTimerPolicy();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.DisableWakeTimers = true;
        _logger.Info("已通过快捷按钮启用软件/定时器唤醒拦截。");
        UpdateSettings(updatedSettings);
    }

    public void RestoreSoftwareWake()
    {
        if (!_settings.DisableWakeTimers)
        {
            _logger.Info("软件/定时器唤醒拦截当前未启用，无需恢复。");
            RefreshWakeTimerPolicySummary();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.DisableWakeTimers = false;
        _logger.Info("已通过快捷按钮请求恢复软件/定时器唤醒策略。");
        UpdateSettings(updatedSettings);
    }

    public void BlockStandbyConnectivityWake()
    {
        if (_settings.DisableStandbyConnectivity)
        {
            _logger.Info("待机联网拦截已启用，正在重新应用当前电源计划策略。");
            ApplyStandbyConnectivityPolicy();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.DisableStandbyConnectivity = true;
        _logger.Info("已通过快捷按钮启用待机联网拦截。");
        UpdateSettings(updatedSettings);
    }

    public void RestoreStandbyConnectivityWake()
    {
        if (!_settings.DisableStandbyConnectivity)
        {
            _logger.Info("待机联网拦截当前未启用，无需恢复。");
            RefreshStandbyConnectivityPolicySummary();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.DisableStandbyConnectivity = false;
        _logger.Info("已通过快捷按钮请求恢复待机联网策略。");
        UpdateSettings(updatedSettings);
    }

    public void EnableBatteryStandbyHibernateFallback()
    {
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            _logger.Info("电池兜底休眠已启用，正在重新应用当前电源计划策略。");
            ApplyBatteryStandbyHibernatePolicy();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.EnforceBatteryStandbyHibernate = true;
        _logger.Info("已通过快捷按钮启用电池待机兜底休眠。");
        UpdateSettings(updatedSettings);
    }

    public void RestoreBatteryStandbyHibernateFallback()
    {
        if (!_settings.EnforceBatteryStandbyHibernate)
        {
            _logger.Info("电池兜底休眠当前未启用，无需恢复。");
            RefreshBatteryStandbyHibernatePolicySummary();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.EnforceBatteryStandbyHibernate = false;
        _logger.Info("已通过快捷按钮请求恢复电池待机休眠策略。");
        UpdateSettings(updatedSettings);
    }

    public void BlockKnownRemoteWakeRequests()
    {
        if (_settings.BlockKnownRemoteWakeRequests)
        {
            _logger.Info("常见远程软件保持唤醒拦截已启用，正在重新应用规则。");
            ApplyKnownRemoteWakePolicy();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.BlockKnownRemoteWakeRequests = true;
        _logger.Info("已通过快捷按钮启用常见远程软件保持唤醒拦截。");
        UpdateSettings(updatedSettings);
    }

    public void RestoreKnownRemoteWakeRequests()
    {
        if (!_settings.BlockKnownRemoteWakeRequests)
        {
            _logger.Info("常见远程软件保持唤醒拦截当前未启用，无需恢复。");
            RefreshKnownRemoteWakePolicySummary();
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        var updatedSettings = _settingsStore.Load();
        updatedSettings.BlockKnownRemoteWakeRequests = false;
        _logger.Info("已通过快捷按钮请求恢复常见远程软件保持唤醒策略。");
        UpdateSettings(updatedSettings);
    }

    public void Dispose()
    {
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        SystemEvents.SessionSwitch -= OnSessionSwitch;
        if (_powerNotificationWindow is not null)
        {
            _powerNotificationWindow.LidStateChanged -= OnLidStateChanged;
            _powerNotificationWindow.Dispose();
        }
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
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var failures = SetWakeTimerPolicy(WakeTimerDisabledValue, WakeTimerDisabledValue);

        if (failures.Length == 0)
        {
            _settings.WakeTimerPolicySummary = snapshot.WakeTimerCaptured
                ? "当前电源计划的唤醒定时器已禁用（AC/DC，可恢复到原始设置）"
                : "当前电源计划的唤醒定时器已禁用（AC/DC，历史基线缺失时将回退到常规启用）";
            _settingsStore.Save(_settings);
            _logger.Info(_settings.WakeTimerPolicySummary);
            return;
        }

        _settings.WakeTimerPolicySummary = "尝试禁用唤醒定时器时收到输出，请检查日志";
        _settingsStore.Save(_settings);
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
            _settings.WakeTimerPolicySummary = hadSnapshot
                ? "已恢复当前电源计划的唤醒定时器到拦截前的 AC/DC 设置"
                : "已恢复当前电源计划的唤醒定时器为常规启用（缺少历史基线）";
            snapshot.WakeTimerCaptured = false;
            snapshot.WakeTimerAcValue = 0;
            snapshot.WakeTimerDcValue = 0;
            _settingsStore.Save(_settings);
            _logger.Info(_settings.WakeTimerPolicySummary);
            return;
        }

        _settings.WakeTimerPolicySummary = hadSnapshot
            ? "尝试恢复唤醒定时器原始设置时收到输出，请检查日志"
            : "尝试恢复唤醒定时器常规启用状态时收到输出，请检查日志";
        _settingsStore.Save(_settings);
        _logger.Warn($"恢复唤醒定时器策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshWakeTimerPolicySummary()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        _settings.WakeTimerPolicySummary = _settings.DisableWakeTimers
            ? (snapshot.WakeTimerCaptured
                ? "已配置为由应用禁用当前电源计划的唤醒定时器，可恢复到原始设置"
                : "已配置为由应用禁用当前电源计划的唤醒定时器；历史基线缺失时将回退到常规启用")
            : "未由应用接管；保留系统当前唤醒定时器策略";
        _settingsStore.Save(_settings);
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
        _settingsStore.Save(_settings);
        _logger.Info($"已记录唤醒定时器原始设置：AC={acValue}，DC={dcValue}。");
    }

    private void ApplyStandbyConnectivityPolicy()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var failures = SetStandbyConnectivityPolicy(
            StandbyConnectivityDisabledValue,
            StandbyConnectivityDisabledValue,
            DisconnectedStandbyModeAggressiveValue,
            DisconnectedStandbyModeAggressiveValue);

        if (failures.Length == 0)
        {
            _settings.StandbyConnectivityPolicySummary = snapshot.StandbyConnectivityCaptured
                ? "当前电源计划已关闭待机联网（AC/DC），并启用主动断网待机模式，可恢复到原始设置"
                : "当前电源计划已关闭待机联网（AC/DC），并启用主动断网待机模式；历史基线缺失时将回退到 Windows 管理/正常模式";
            _settingsStore.Save(_settings);
            _logger.Info(_settings.StandbyConnectivityPolicySummary);
            return;
        }

        _settings.StandbyConnectivityPolicySummary = "尝试关闭待机联网时收到输出，请检查日志";
        _settingsStore.Save(_settings);
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
            _settings.StandbyConnectivityPolicySummary = hadSnapshot
                ? "已恢复当前电源计划的待机联网和断网待机模式到拦截前的 AC/DC 设置"
                : "已恢复当前电源计划的待机联网为 Windows 管理，断网待机模式为正常（缺少历史基线）";
            snapshot.StandbyConnectivityCaptured = false;
            snapshot.StandbyConnectivityAcValue = 0;
            snapshot.StandbyConnectivityDcValue = 0;
            snapshot.DisconnectedStandbyModeAcValue = 0;
            snapshot.DisconnectedStandbyModeDcValue = 0;
            _settingsStore.Save(_settings);
            _logger.Info(_settings.StandbyConnectivityPolicySummary);
            return;
        }

        _settings.StandbyConnectivityPolicySummary = hadSnapshot
            ? "尝试恢复待机联网原始设置时收到输出，请检查日志"
            : "尝试恢复待机联网默认设置时收到输出，请检查日志";
        _settingsStore.Save(_settings);
        _logger.Warn($"恢复待机联网策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshStandbyConnectivityPolicySummary()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        if (_settings.DisableStandbyConnectivity)
        {
            _settings.StandbyConnectivityPolicySummary = snapshot.StandbyConnectivityCaptured
                ? "已配置为关闭待机联网并启用主动断网待机模式，可恢复到原始设置"
                : "已配置为关闭待机联网并启用主动断网待机模式；历史基线缺失时将回退到 Windows 管理/正常模式";
            _settingsStore.Save(_settings);
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

        _settingsStore.Save(_settings);
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
        _settingsStore.Save(_settings);
        _logger.Info($"已记录待机联网原始设置：联网 AC={connectivityAcValue}，DC={connectivityDcValue}；断网待机 AC={disconnectedAcValue}，DC={disconnectedDcValue}。");
    }

    private void ApplyBatteryStandbyHibernatePolicy()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var timeoutSeconds = GetBatteryStandbyHibernateTimeoutSeconds();
        var acValue = snapshot.BatteryStandbyHibernateCaptured
            ? snapshot.BatteryStandbyHibernateAcValue
            : TryGetCurrentHibernateAfterStandbyIndices(out var currentAcValue, out _)
                ? currentAcValue
                : HibernateAfterStandbyDisabledValue;
        var failures = SetHibernateAfterStandbyPolicy(acValue, timeoutSeconds);

        if (failures.Length == 0)
        {
            _settings.BatteryStandbyHibernatePolicySummary = snapshot.BatteryStandbyHibernateCaptured
                ? $"当前电源计划已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠（DC，可恢复到原始设置）"
                : $"当前电源计划已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠（DC，历史基线缺失时将回退到永不休眠）";
            _settingsStore.Save(_settings);
            _logger.Info(_settings.BatteryStandbyHibernatePolicySummary);
            return;
        }

        _settings.BatteryStandbyHibernatePolicySummary = "尝试配置电池待机兜底休眠时收到输出，请检查日志";
        _settingsStore.Save(_settings);
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
            _settings.BatteryStandbyHibernatePolicySummary = hadSnapshot
                ? "已恢复当前电源计划的待机后休眠时间到拦截前的 AC/DC 设置"
                : "已恢复当前电源计划的电池待机后休眠为永不（缺少历史基线）";
            snapshot.BatteryStandbyHibernateCaptured = false;
            snapshot.BatteryStandbyHibernateAcValue = 0;
            snapshot.BatteryStandbyHibernateDcValue = 0;
            _settingsStore.Save(_settings);
            _logger.Info(_settings.BatteryStandbyHibernatePolicySummary);
            return;
        }

        _settings.BatteryStandbyHibernatePolicySummary = hadSnapshot
            ? "尝试恢复待机后休眠原始设置时收到输出，请检查日志"
            : "尝试恢复电池待机后休眠默认设置时收到输出，请检查日志";
        _settingsStore.Save(_settings);
        _logger.Warn($"恢复电池待机休眠策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshBatteryStandbyHibernatePolicySummary()
    {
        var snapshot = GetOrCreateActivePowerPlanSnapshot();
        var timeoutSeconds = GetBatteryStandbyHibernateTimeoutSeconds();
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            _settings.BatteryStandbyHibernatePolicySummary = snapshot.BatteryStandbyHibernateCaptured
                ? $"已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠（DC），可恢复到原始设置"
                : $"已配置为电池待机 {DescribePowerSettingDuration(timeoutSeconds)}后自动休眠（DC）；历史基线缺失时将回退到永不休眠";
            _settingsStore.Save(_settings);
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

        _settingsStore.Save(_settings);
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
        _settingsStore.Save(_settings);
        _logger.Info($"已记录待机后休眠原始设置：AC={acValue}，DC={dcValue}。");
    }

    private void ApplyKnownRemoteWakePolicy()
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);
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

        if (failures.Count == 0)
        {
            _settings.KnownRemoteWakePolicySummary =
                _settings.KnownRemoteWakeRequestBackupCaptured
                    ? $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {appliedCount}/{managedEntries.Count} 条规则，可恢复到拦截前状态）"
                    : $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {appliedCount}/{managedEntries.Count} 条规则，历史基线缺失时将按清除处理）";
            _settingsStore.Save(_settings);
            _logger.Info($"{_settings.KnownRemoteWakePolicySummary}。覆盖：{RemoteWakeBlockCatalog.ProductSummary(_settings.CustomRemoteWakeEntries)}。");
            return;
        }

        _settings.KnownRemoteWakePolicySummary = $"应用远程软件拦截规则时收到输出，请检查日志（已命中 {appliedCount}/{managedEntries.Count} 条规则）";
        _settingsStore.Save(_settings);
        _logger.Warn($"应用远程软件拦截规则时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RestoreKnownRemoteWakePolicy()
    {
        var failures = new List<string>();
        var hadSnapshot = _settings.KnownRemoteWakeRequestBackupCaptured;
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);

        foreach (var entry in managedEntries)
        {
            var backupKey = BuildRequestOverrideBackupKey(entry);
            if (hadSnapshot
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

        if (failures.Count == 0)
        {
            _settings.KnownRemoteWakePolicySummary = hadSnapshot
                ? "已恢复常见远程软件拦截前的系统请求替代状态"
                : "已清除应用管理的常见远程软件拦截规则（缺少历史基线时按清除处理）";
            _settings.KnownRemoteWakeRequestBackupCaptured = false;
            _settings.KnownRemoteWakeRequestOverrideBackup = [];
            _settingsStore.Save(_settings);
            _logger.Info(_settings.KnownRemoteWakePolicySummary);
            return;
        }

        _settings.KnownRemoteWakePolicySummary = hadSnapshot
            ? "尝试恢复常见远程软件拦截前状态时收到输出，请检查日志"
            : "尝试清除常见远程软件拦截规则时收到输出，请检查日志";
        _settingsStore.Save(_settings);
        _logger.Warn($"恢复远程软件拦截规则时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshKnownRemoteWakePolicySummary()
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);
        var managedCount = CountKnownRemoteRequestOverrides(managedEntries, static value => value > 0);

        _settings.KnownRemoteWakePolicySummary = _settings.BlockKnownRemoteWakeRequests
            ? (_settings.KnownRemoteWakeRequestBackupCaptured
                ? $"已配置为拦截常见远程软件的保持唤醒请求（当前检测到 {managedCount} 条规则，可恢复到拦截前状态）"
                : $"已配置为拦截常见远程软件的保持唤醒请求（当前检测到 {managedCount} 条规则，历史基线缺失时将按清除处理）")
            : managedCount == 0
                ? "未由应用接管；未检测到常见远程软件保持唤醒拦截规则"
                : $"未由应用接管；系统当前仍存在 {managedCount} 条常见远程软件请求替代";
        _settingsStore.Save(_settings);
    }

    private void CaptureKnownRemoteWakeRestoreSnapshotIfNeeded()
    {
        if (_settings.KnownRemoteWakeRequestBackupCaptured)
        {
            return;
        }

        var snapshot = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries))
        {
            if (!TryGetRequestOverrideValue(entry, out var value))
            {
                continue;
            }

            snapshot[BuildRequestOverrideBackupKey(entry)] = value;
        }

        _settings.KnownRemoteWakeRequestOverrideBackup = snapshot;
        _settings.KnownRemoteWakeRequestBackupCaptured = true;
        _settingsStore.Save(_settings);
        _logger.Info($"已记录常见远程软件请求替代基线（现有规则 {snapshot.Count} 条）。");
    }

    private void EnsureAutostartMatchesSettings()
    {
        if (AutostartManager.IsEnabled() != _settings.StartWithWindows)
        {
            AutostartManager.SetEnabled(_settings.StartWithWindows);
            _logger.Info(_settings.StartWithWindows ? "已启用开机自启。" : "已关闭开机自启。");
        }
    }

    private void OnPowerModeChanged(object sender, PowerModeChangedEventArgs e)
    {
        switch (e.Mode)
        {
            case PowerModes.Suspend:
                CancelPendingResumeProtection();
                ClearManualResumeSignals();
                _settings.LastSuspendUtc = DateTimeOffset.UtcNow;
                _settingsStore.Save(_settings);
                _logger.Warn("系统即将进入挂起/休眠。");
                StateChanged?.Invoke(this, EventArgs.Empty);
                break;

            case PowerModes.Resume:
                _settings.LastResumeUtc = DateTimeOffset.UtcNow;
                var snapshot = CollectWakeDiagnosticSnapshot(includePowerRequests: false, includeSleepStudy: false);
                var analysis = AnalyzeWake(snapshot);
                if (TryGetRecentManualResumeSignal(out var manualSignalReason) && analysis.Kind != WakeKind.Manual)
                {
                    analysis = new WakeAnalysis(WakeKind.Manual, $"检测到人工行为（{manualSignalReason}）");
                }
                _settings.LastWakeSummary = analysis.Summary;
                _settings.LastWakeEvidenceSummary = _wakeDiagnostics.SummarizeEvidence(snapshot);
                _settingsStore.Save(_settings);
                _logger.Warn($"系统已从挂起/休眠恢复。判定：{analysis.Summary}{Environment.NewLine}{BuildWakeDiagnosticsText(snapshot, includePowerRequests: false, includeSleepStudy: false)}");
                StateChanged?.Invoke(this, EventArgs.Empty);
                ArmResumeProtectionIfNeeded(analysis);
                break;

            case PowerModes.StatusChange:
                StateChanged?.Invoke(this, EventArgs.Empty);
                break;
        }
    }

    private void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
    {
        if (e.Reason is not (SessionSwitchReason.SessionUnlock
            or SessionSwitchReason.ConsoleConnect
            or SessionSwitchReason.RemoteConnect
            or SessionSwitchReason.SessionLogon))
        {
            return;
        }

        var reason = DescribeSessionSwitchReason(e.Reason);
        RecordManualResumeSignal(reason);

        if (CancelPendingResumeProtection($"检测到{reason}，已取消待执行的自动回睡。"))
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private void OnLidStateChanged(object? sender, bool isOpen)
    {
        if (!isOpen)
        {
            return;
        }

        RecordManualResumeSignal("开盖");

        if (CancelPendingResumeProtection("检测到开盖操作，已取消待执行的自动回睡。"))
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private void ArmResumeProtectionIfNeeded(WakeAnalysis analysis)
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
        lock (_resumeProtectionSync)
        {
            _resumeProtectionArmedAtTick = NativeMethods.GetTickCount();
            _resumeTimer.Change(checked(delaySeconds * 1000), Timeout.Infinite);
        }
        _logger.Warn($"已启动恢复保护，将在 {delaySeconds} 秒后自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
    }

    private void OnResumeTimerElapsed()
    {
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
            _logger.Info("检测到恢复后已有人工操作，已取消本次自动回睡。");
            StateChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        if (_settings.ResumeProtectionMode == ResumeProtectionMode.Hibernate)
        {
            RequestSuspend(PowerState.Hibernate, "恢复保护到期，执行自动休眠。");
        }
        else
        {
            RequestSuspend(PowerState.Suspend, "恢复保护到期，执行自动睡眠。");
        }
    }

    private static string DescribeMode(PowerPolicyMode mode)
    {
        return mode == PowerPolicyMode.KeepAwakeIndefinitely ? "无限期保持激活" : "遵循电源计划";
    }

    private static string DescribeResumeProtection(ResumeProtectionMode mode)
    {
        return mode == ResumeProtectionMode.Hibernate ? "休眠" : "睡眠";
    }

    private string BuildProtectionRuleSummary()
    {
        if (_settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            return "当前处于无限保持唤醒模式，不执行恢复后自动回睡。";
        }

        if (!_settings.ResumeProtectionEnabled)
        {
            return "恢复保护已关闭，系统恢复后不会自动重新睡眠/休眠。";
        }

        var resumeAction = DescribeResumeProtection(_settings.ResumeProtectionMode);
        var delaySeconds = Math.Max(3, _settings.ResumeProtectionDelaySeconds);

        if (_settings.ResumeProtectionOnlyForUnattendedWake)
        {
            return $"人工行为（键盘、鼠标、开盖、解锁、控制台/远程接管、登录）恢复后跳过自动回睡；其他恢复（软件、定时器、设备、来源不明）会在 {delaySeconds} 秒后自动{resumeAction}。";
        }

        return $"不区分人工或非人工，系统每次恢复后都会在 {delaySeconds} 秒后自动{resumeAction}。";
    }

    private string RunPowerCfg(string arguments)
    {
        return _powerCfg.Run(arguments);
    }

    private static WakeAnalysis AnalyzeWake(WakeDiagnosticSnapshot snapshot)
    {
        var lastWake = snapshot.LastWakeText.ToLowerInvariant();
        var wakeTimers = snapshot.WakeTimersText.ToLowerInvariant();
        var eventSummary = snapshot.EventSummary.ToLowerInvariant();
        var combined = string.Join(
            Environment.NewLine,
            snapshot.LastWakeText,
            snapshot.WakeTimersText,
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
        _settings.PowerPlanRestoreSnapshots ??= new Dictionary<string, PowerPlanRestoreSnapshot>(StringComparer.OrdinalIgnoreCase);
        _settings.CustomRemoteWakeEntries = RemoteWakeBlockCatalog.NormalizeCustomEntries(_settings.CustomRemoteWakeEntries).ToList();
        _settings.BatteryStandbyHibernateTimeoutSeconds = Math.Clamp(
            _settings.BatteryStandbyHibernateTimeoutSeconds <= 0
                ? DefaultBatteryStandbyHibernateTimeoutSeconds
                : _settings.BatteryStandbyHibernateTimeoutSeconds,
            180,
            14400);
        _settings.LastWakeSummary ??= "无";
        _settings.LastWakeEvidenceSummary ??= "无";
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

    private int GetBatteryStandbyHibernateTimeoutSeconds()
    {
        _settings.BatteryStandbyHibernateTimeoutSeconds = Math.Clamp(
            _settings.BatteryStandbyHibernateTimeoutSeconds <= 0
                ? DefaultBatteryStandbyHibernateTimeoutSeconds
                : _settings.BatteryStandbyHibernateTimeoutSeconds,
            180,
            14400);
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

    private string BuildCapabilitySummary()
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
            parts.Add("远控拦截、请求替代和部分 powercfg 写入可能受限");
        }

        if (!string.IsNullOrWhiteSpace(_settings.KnownRemoteWakePolicySummary)
            && _settings.BlockKnownRemoteWakeRequests
            && _settings.KnownRemoteWakePolicySummary.Contains("请检查日志", StringComparison.OrdinalIgnoreCase))
        {
            parts.Add("远控拦截最近一次应用未完全成功");
        }

        return string.Join("；", parts) + "。";
    }

    private string BuildRiskSummary()
    {
        var risks = new List<string>();

        if (_settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            risks.Add("当前处于无限保持唤醒模式");
        }
        else
        {
            if (!_settings.ResumeProtectionEnabled)
            {
                risks.Add("恢复保护已关闭");
            }

            if (!_settings.DisableWakeTimers)
            {
                risks.Add("唤醒定时器未接管");
            }

            if (!_settings.DisableStandbyConnectivity)
            {
                risks.Add("待机联网未接管");
            }

            if (!_settings.EnforceBatteryStandbyHibernate)
            {
                risks.Add("电池兜底休眠未接管");
            }
        }

        if (!_settings.BlockKnownRemoteWakeRequests)
        {
            risks.Add("远控保活拦截未启用");
        }

        return risks.Count == 0
            ? "当前保护层完整，没有明显高风险缺口。"
            : $"当前最主要风险：{string.Join("；", risks.Take(3))}。";
    }

    private string BuildPowerPlanSummary()
    {
        return TryGetActivePowerPlanInfo(out var planId, out var planName)
            ? $"当前电源计划：{planName} ({planId})"
            : "当前电源计划：未能识别 GUID";
    }

    private string BuildManagedRemoteEntriesSummary()
    {
        var managedEntries = RemoteWakeBlockCatalog.GetManagedEntries(_settings.CustomRemoteWakeEntries);
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
        _logger.Warn(logMessage);
        if (Application.SetSuspendState(powerState, false, false))
        {
            return;
        }

        _logger.Error($"请求{DescribePowerState(powerState)}失败，系统拒绝执行。");
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

    private enum WakeKind
    {
        Manual,
        Unattended,
        Unknown
    }

    private readonly record struct WakeAnalysis(WakeKind Kind, string Summary);
}
