using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class PowerController : IDisposable
{
    private static readonly TimeSpan ManualResumeSignalWindow = TimeSpan.FromSeconds(20);
    private const int PowerCfgTimeoutMilliseconds = 3000;
    private const int WakeTimerDisabledValue = 0;
    private const int WakeTimerEnabledValue = 1;
    private const int StandbyConnectivityDisabledValue = 0;
    private const int StandbyConnectivityEnabledValue = 1;
    private const int StandbyConnectivityManagedByWindowsValue = 2;
    private const int DisconnectedStandbyModeNormalValue = 0;
    private const int DisconnectedStandbyModeAggressiveValue = 1;
    private const int HibernateAfterStandbyDisabledValue = 0;
    private const int HibernateAfterStandbyFallbackRestoreDcValue = int.MaxValue;
    private const int BatteryStandbyHibernateTimeoutSeconds = 600;
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
            if (!_settings.BatteryStandbyHibernateRestoreSnapshotCaptured)
            {
                CaptureBatteryStandbyHibernateRestoreSnapshotIfNeeded();
            }

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
        _logger.Info($"设置已更新：模式={DescribeMode(_settings.PolicyMode)}，恢复保护={_settings.ResumeProtectionEnabled}，待机联网拦截={_settings.DisableStandbyConnectivity}，电池兜底休眠={_settings.EnforceBatteryStandbyHibernate}，远控拦截={_settings.BlockKnownRemoteWakeRequests}。");
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
        var lastWake = RunPowerCfg("/lastwake");
        var wakeTimers = RunPowerCfg("/waketimers");
        return $"lastwake:{Environment.NewLine}{lastWake}{Environment.NewLine}{Environment.NewLine}waketimers:{Environment.NewLine}{wakeTimers}";
    }

    public string CollectPowerRequestDiagnostics()
    {
        var requests = RunPowerCfg("/requests");
        var overrides = RunPowerCfg("/requestsoverride");
        return $"requests:{Environment.NewLine}{requests}{Environment.NewLine}{Environment.NewLine}requestsoverride:{Environment.NewLine}{overrides}";
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
        var failures = SetWakeTimerPolicy(WakeTimerDisabledValue, WakeTimerDisabledValue);

        if (failures.Length == 0)
        {
            _settings.WakeTimerPolicySummary = _settings.WakeTimerRestoreSnapshotCaptured
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
        var hadSnapshot = _settings.WakeTimerRestoreSnapshotCaptured;
        var acValue = hadSnapshot ? _settings.WakeTimerRestoreAcValue : WakeTimerEnabledValue;
        var dcValue = hadSnapshot ? _settings.WakeTimerRestoreDcValue : WakeTimerEnabledValue;
        var failures = SetWakeTimerPolicy(acValue, dcValue);

        if (failures.Length == 0)
        {
            _settings.WakeTimerPolicySummary = hadSnapshot
                ? "已恢复当前电源计划的唤醒定时器到拦截前的 AC/DC 设置"
                : "已恢复当前电源计划的唤醒定时器为常规启用（缺少历史基线）";
            ClearWakeTimerRestoreSnapshot();
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
        _settings.WakeTimerPolicySummary = _settings.DisableWakeTimers
            ? (_settings.WakeTimerRestoreSnapshotCaptured
                ? "已配置为由应用禁用当前电源计划的唤醒定时器，可恢复到原始设置"
                : "已配置为由应用禁用当前电源计划的唤醒定时器；历史基线缺失时将回退到常规启用")
            : "未由应用接管；保留系统当前唤醒定时器策略";
        _settingsStore.Save(_settings);
    }

    private void CaptureWakeTimerRestoreSnapshotIfNeeded()
    {
        if (_settings.WakeTimerRestoreSnapshotCaptured)
        {
            return;
        }

        if (!TryGetCurrentWakeTimerIndices(out var acValue, out var dcValue))
        {
            _logger.Warn("未能记录唤醒定时器原始 AC/DC 设置；恢复时将回退到常规启用。");
            return;
        }

        _settings.WakeTimerRestoreAcValue = acValue;
        _settings.WakeTimerRestoreDcValue = dcValue;
        _settings.WakeTimerRestoreSnapshotCaptured = true;
        _settingsStore.Save(_settings);
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
            _settings.StandbyConnectivityPolicySummary = _settings.StandbyConnectivityRestoreSnapshotCaptured
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
        var hadSnapshot = _settings.StandbyConnectivityRestoreSnapshotCaptured;
        var connectivityAcValue = hadSnapshot ? _settings.StandbyConnectivityRestoreAcValue : StandbyConnectivityManagedByWindowsValue;
        var connectivityDcValue = hadSnapshot ? _settings.StandbyConnectivityRestoreDcValue : StandbyConnectivityManagedByWindowsValue;
        var disconnectedAcValue = hadSnapshot ? _settings.DisconnectedStandbyModeRestoreAcValue : DisconnectedStandbyModeNormalValue;
        var disconnectedDcValue = hadSnapshot ? _settings.DisconnectedStandbyModeRestoreDcValue : DisconnectedStandbyModeNormalValue;
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
            ClearStandbyConnectivityRestoreSnapshot();
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
        if (_settings.DisableStandbyConnectivity)
        {
            _settings.StandbyConnectivityPolicySummary = _settings.StandbyConnectivityRestoreSnapshotCaptured
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
        if (_settings.StandbyConnectivityRestoreSnapshotCaptured)
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

        _settings.StandbyConnectivityRestoreAcValue = connectivityAcValue;
        _settings.StandbyConnectivityRestoreDcValue = connectivityDcValue;
        _settings.DisconnectedStandbyModeRestoreAcValue = disconnectedAcValue;
        _settings.DisconnectedStandbyModeRestoreDcValue = disconnectedDcValue;
        _settings.StandbyConnectivityRestoreSnapshotCaptured = true;
        _settingsStore.Save(_settings);
        _logger.Info($"已记录待机联网原始设置：联网 AC={connectivityAcValue}，DC={connectivityDcValue}；断网待机 AC={disconnectedAcValue}，DC={disconnectedDcValue}。");
    }

    private void ApplyBatteryStandbyHibernatePolicy()
    {
        var acValue = _settings.BatteryStandbyHibernateRestoreSnapshotCaptured
            ? _settings.BatteryStandbyHibernateRestoreAcValue
            : TryGetCurrentHibernateAfterStandbyIndices(out var currentAcValue, out _)
                ? currentAcValue
                : HibernateAfterStandbyDisabledValue;
        var failures = SetHibernateAfterStandbyPolicy(acValue, BatteryStandbyHibernateTimeoutSeconds);

        if (failures.Length == 0)
        {
            _settings.BatteryStandbyHibernatePolicySummary = _settings.BatteryStandbyHibernateRestoreSnapshotCaptured
                ? $"当前电源计划已配置为电池待机 {DescribePowerSettingDuration(BatteryStandbyHibernateTimeoutSeconds)}后自动休眠（DC，可恢复到原始设置）"
                : $"当前电源计划已配置为电池待机 {DescribePowerSettingDuration(BatteryStandbyHibernateTimeoutSeconds)}后自动休眠（DC，历史基线缺失时将回退到永不休眠）";
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
        var hadSnapshot = _settings.BatteryStandbyHibernateRestoreSnapshotCaptured;
        var acValue = hadSnapshot
            ? _settings.BatteryStandbyHibernateRestoreAcValue
            : TryGetCurrentHibernateAfterStandbyIndices(out var currentAcValue, out _)
                ? currentAcValue
                : HibernateAfterStandbyDisabledValue;
        var dcValue = hadSnapshot ? _settings.BatteryStandbyHibernateRestoreDcValue : HibernateAfterStandbyFallbackRestoreDcValue;
        var failures = SetHibernateAfterStandbyPolicy(acValue, dcValue);

        if (failures.Length == 0)
        {
            _settings.BatteryStandbyHibernatePolicySummary = hadSnapshot
                ? "已恢复当前电源计划的待机后休眠时间到拦截前的 AC/DC 设置"
                : "已恢复当前电源计划的电池待机后休眠为永不（缺少历史基线）";
            ClearBatteryStandbyHibernateRestoreSnapshot();
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
        if (_settings.EnforceBatteryStandbyHibernate)
        {
            _settings.BatteryStandbyHibernatePolicySummary = _settings.BatteryStandbyHibernateRestoreSnapshotCaptured
                ? $"已配置为电池待机 {DescribePowerSettingDuration(BatteryStandbyHibernateTimeoutSeconds)}后自动休眠（DC），可恢复到原始设置"
                : $"已配置为电池待机 {DescribePowerSettingDuration(BatteryStandbyHibernateTimeoutSeconds)}后自动休眠（DC）；历史基线缺失时将回退到永不休眠";
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
        if (_settings.BatteryStandbyHibernateRestoreSnapshotCaptured)
        {
            return;
        }

        if (!TryGetCurrentHibernateAfterStandbyIndices(out var acValue, out var dcValue))
        {
            _logger.Warn("未能记录待机后休眠原始 AC/DC 设置；恢复时将回退到电池永不休眠。");
            return;
        }

        _settings.BatteryStandbyHibernateRestoreAcValue = acValue;
        _settings.BatteryStandbyHibernateRestoreDcValue = dcValue;
        _settings.BatteryStandbyHibernateRestoreSnapshotCaptured = true;
        _settingsStore.Save(_settings);
        _logger.Info($"已记录待机后休眠原始设置：AC={acValue}，DC={dcValue}。");
    }

    private void ApplyKnownRemoteWakePolicy()
    {
        var failures = new List<string>();

        foreach (var entry in RemoteWakeBlockCatalog.Entries)
        {
            var result = ApplyRequestOverride(entry, FullRequestOverrideMask);
            if (!string.IsNullOrWhiteSpace(result))
            {
                failures.Add($"{entry.CallerTypeArgument}:{entry.Name}: {result}");
            }
        }

        var appliedCount = CountKnownRemoteRequestOverrides(static value => value == FullRequestOverrideMask);

        if (failures.Count == 0)
        {
            _settings.KnownRemoteWakePolicySummary =
                _settings.KnownRemoteWakeRequestBackupCaptured
                    ? $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {appliedCount}/{RemoteWakeBlockCatalog.Entries.Count} 条规则，可恢复到拦截前状态）"
                    : $"已拦截常见远程软件的 DISPLAY/SYSTEM/AWAYMODE 保持唤醒请求（命中 {appliedCount}/{RemoteWakeBlockCatalog.Entries.Count} 条规则，历史基线缺失时将按清除处理）";
            _settingsStore.Save(_settings);
            _logger.Info($"{_settings.KnownRemoteWakePolicySummary}。覆盖：{RemoteWakeBlockCatalog.ProductSummary}。");
            return;
        }

        _settings.KnownRemoteWakePolicySummary = $"应用远程软件拦截规则时收到输出，请检查日志（已命中 {appliedCount}/{RemoteWakeBlockCatalog.Entries.Count} 条规则）";
        _settingsStore.Save(_settings);
        _logger.Warn($"应用远程软件拦截规则时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RestoreKnownRemoteWakePolicy()
    {
        var failures = new List<string>();
        var hadSnapshot = _settings.KnownRemoteWakeRequestBackupCaptured;

        foreach (var entry in RemoteWakeBlockCatalog.Entries)
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
            ClearKnownRemoteWakeRestoreSnapshot();
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
        var managedCount = CountKnownRemoteRequestOverrides(static value => value > 0);

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
        foreach (var entry in RemoteWakeBlockCatalog.Entries)
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
                var diagnostics = CollectWakeDiagnostics();
                var analysis = AnalyzeWake(diagnostics);
                if (TryGetRecentManualResumeSignal(out var manualSignalReason) && analysis.Kind != WakeKind.Manual)
                {
                    analysis = new WakeAnalysis(WakeKind.Manual, $"检测到人工行为（{manualSignalReason}）");
                }
                _settings.LastWakeSummary = analysis.Summary;
                _settingsStore.Save(_settings);
                _logger.Warn($"系统已从挂起/休眠恢复。判定：{analysis.Summary}{Environment.NewLine}{diagnostics}");
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
        try
        {
            using var process = new Process();
            process.StartInfo.FileName = "powercfg";
            process.StartInfo.Arguments = arguments;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            if (!process.WaitForExit(PowerCfgTimeoutMilliseconds))
            {
                TryKillProcess(process);
                process.WaitForExit();
                var timeoutMessage = $"powercfg {arguments} 超时（超过 {PowerCfgTimeoutMilliseconds / 1000} 秒）";
                _logger.Warn(timeoutMessage);
                return timeoutMessage;
            }

            Task.WaitAll(new Task[] { outputTask, errorTask }, 1000);
            var output = outputTask.IsCompletedSuccessfully ? outputTask.Result : string.Empty;
            var error = errorTask.IsCompletedSuccessfully ? errorTask.Result : string.Empty;
            if (process.ExitCode == 0)
            {
                return string.IsNullOrWhiteSpace(output) ? string.Empty : output.Trim();
            }

            var message = string.IsNullOrWhiteSpace(error) ? output.Trim() : error.Trim();
            return string.IsNullOrWhiteSpace(message) ? $"powercfg {arguments} 失败，退出码 {process.ExitCode}" : message;
        }
        catch (Exception ex)
        {
            _logger.Error($"执行 powercfg {arguments} 失败：{ex.Message}");
            return $"powercfg {arguments} 调用失败：{ex.Message}";
        }
    }

    private static WakeAnalysis AnalyzeWake(string diagnostics)
    {
        var (lastWakeText, wakeTimersText) = SplitDiagnostics(diagnostics);
        var lastWake = lastWakeText.ToLowerInvariant();
        var wakeTimers = wakeTimersText.ToLowerInvariant();

        if (ContainsAny(lastWake, ManualWakeIndicators))
        {
            return new WakeAnalysis(WakeKind.Manual, "疑似人工唤醒（按键/开盖/鼠标键盘）");
        }

        if (ContainsAny(lastWake, DeviceWakeIndicators))
        {
            return new WakeAnalysis(WakeKind.Unattended, "疑似设备或网络唤醒");
        }

        if (ContainsAny(lastWake, TimerWakeIndicators)
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

    private static (string LastWake, string WakeTimers) SplitDiagnostics(string diagnostics)
    {
        const string lastWakeLabel = "lastwake:";
        const string wakeTimersLabel = "waketimers:";

        var wakeTimersIndex = diagnostics.IndexOf(wakeTimersLabel, StringComparison.OrdinalIgnoreCase);
        var lastWake = diagnostics;
        var wakeTimers = string.Empty;

        if (wakeTimersIndex >= 0)
        {
            lastWake = diagnostics[..wakeTimersIndex];
            wakeTimers = diagnostics[(wakeTimersIndex + wakeTimersLabel.Length)..];
        }

        lastWake = lastWake.Replace(lastWakeLabel, string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        wakeTimers = wakeTimers.Trim();
        return (lastWake, wakeTimers);
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

    private int CountKnownRemoteRequestOverrides(Predicate<int> predicate)
    {
        var count = 0;
        foreach (var entry in RemoteWakeBlockCatalog.Entries)
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

    private void ClearWakeTimerRestoreSnapshot()
    {
        _settings.WakeTimerRestoreSnapshotCaptured = false;
        _settings.WakeTimerRestoreAcValue = 0;
        _settings.WakeTimerRestoreDcValue = 0;
    }

    private void ClearStandbyConnectivityRestoreSnapshot()
    {
        _settings.StandbyConnectivityRestoreSnapshotCaptured = false;
        _settings.StandbyConnectivityRestoreAcValue = 0;
        _settings.StandbyConnectivityRestoreDcValue = 0;
        _settings.DisconnectedStandbyModeRestoreAcValue = 0;
        _settings.DisconnectedStandbyModeRestoreDcValue = 0;
    }

    private void ClearBatteryStandbyHibernateRestoreSnapshot()
    {
        _settings.BatteryStandbyHibernateRestoreSnapshotCaptured = false;
        _settings.BatteryStandbyHibernateRestoreAcValue = 0;
        _settings.BatteryStandbyHibernateRestoreDcValue = 0;
    }

    private void ClearKnownRemoteWakeRestoreSnapshot()
    {
        _settings.KnownRemoteWakeRequestBackupCaptured = false;
        _settings.KnownRemoteWakeRequestOverrideBackup = [];
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

    private enum WakeKind
    {
        Manual,
        Unattended,
        Unknown
    }

    private readonly record struct WakeAnalysis(WakeKind Kind, string Summary);
}
