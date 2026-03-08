using System.Diagnostics;
using Microsoft.Win32;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class PowerController : IDisposable
{
    private readonly SettingsStore _settingsStore;
    private readonly FileLogger _logger;
    private readonly System.Windows.Forms.Timer _resumeTimer;
    private AppSettings _settings;

    public PowerController(SettingsStore settingsStore, FileLogger logger, AppSettings settings)
    {
        _settingsStore = settingsStore;
        _logger = logger;
        _settings = settings;
        _resumeTimer = new System.Windows.Forms.Timer();
        _resumeTimer.Tick += ResumeTimerOnTick;

        SystemEvents.PowerModeChanged += OnPowerModeChanged;
        ApplyPolicy(_settings.PolicyMode);
        RefreshWakeTimerPolicySummary();
        EnsureAutostartMatchesSettings();
        _logger.Info($"应用启动，当前模式：{DescribeMode(_settings.PolicyMode)}。");
    }

    public event EventHandler? StateChanged;

    public AppSettings CurrentSettings => _settings;

    public string CurrentStatus =>
        _settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely
            ? "无限保持唤醒"
            : (_settings.ResumeProtectionEnabled
                ? $"遵循电源计划，恢复后 {_settings.ResumeProtectionDelaySeconds} 秒自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}"
                : "遵循电源计划");

    public void UpdateSettings(AppSettings updatedSettings)
    {
        _settings = updatedSettings;
        _settingsStore.Save(_settings);
        ApplyPolicy(_settings.PolicyMode);
        if (_settings.DisableWakeTimers)
        {
            ApplyWakeTimerPolicy();
        }
        else
        {
            RefreshWakeTimerPolicySummary();
        }
        EnsureAutostartMatchesSettings();
        _logger.Info($"设置已更新：模式={DescribeMode(_settings.PolicyMode)}，恢复保护={_settings.ResumeProtectionEnabled}。");
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SleepNow()
    {
        _logger.Warn("用户手动请求立即睡眠。");
        Application.SetSuspendState(PowerState.Suspend, false, false);
    }

    public void HibernateNow()
    {
        _logger.Warn("用户手动请求立即休眠。");
        Application.SetSuspendState(PowerState.Hibernate, false, false);
    }

    public string CollectWakeDiagnostics()
    {
        var lastWake = RunPowerCfg("/lastwake");
        var wakeTimers = RunPowerCfg("/waketimers");
        return $"lastwake:{Environment.NewLine}{lastWake}{Environment.NewLine}{Environment.NewLine}waketimers:{Environment.NewLine}{wakeTimers}";
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

    public void Dispose()
    {
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        _resumeTimer.Tick -= ResumeTimerOnTick;
        _resumeTimer.Dispose();
        NativeMethods.SetThreadExecutionState(ExecutionState.EsContinuous);
    }

    private void ApplyPolicy(PowerPolicyMode mode)
    {
        if (mode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            NativeMethods.SetThreadExecutionState(ExecutionState.EsContinuous | ExecutionState.EsSystemRequired);
            _logger.Info("已启用无限保持唤醒。");
        }
        else
        {
            NativeMethods.SetThreadExecutionState(ExecutionState.EsContinuous);
            _logger.Info("已切换为遵循电源计划。");
        }
    }

    private void ApplyWakeTimerPolicy()
    {
        const int desiredIndex = 0;
        const string actionText = "禁用";

        var acResult = RunPowerCfg($"/setacvalueindex scheme_current sub_sleep rtcwake {desiredIndex}");
        var dcResult = RunPowerCfg($"/setdcvalueindex scheme_current sub_sleep rtcwake {desiredIndex}");
        var activateResult = RunPowerCfg("/setactive scheme_current");

        var failures = new[] { acResult, dcResult, activateResult }
            .Where(static x => !string.IsNullOrWhiteSpace(x))
            .ToArray();

        if (failures.Length == 0)
        {
            _settings.WakeTimerPolicySummary = $"当前电源计划的唤醒定时器已{actionText}（AC/DC）";
            _settingsStore.Save(_settings);
            _logger.Info(_settings.WakeTimerPolicySummary);
            return;
        }

        _settings.WakeTimerPolicySummary = $"尝试{actionText}唤醒定时器时收到输出，请检查日志";
        _settingsStore.Save(_settings);
        _logger.Warn($"切换唤醒定时器策略时收到输出：{Environment.NewLine}{string.Join(Environment.NewLine, failures)}");
    }

    private void RefreshWakeTimerPolicySummary()
    {
        _settings.WakeTimerPolicySummary = _settings.DisableWakeTimers
            ? "已配置为由应用禁用当前电源计划的唤醒定时器"
            : "未由应用接管；保留系统当前唤醒定时器策略";
        _settingsStore.Save(_settings);
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
                _settings.LastSuspendUtc = DateTimeOffset.UtcNow;
                _settingsStore.Save(_settings);
                _logger.Warn("系统即将进入挂起/休眠。");
                StateChanged?.Invoke(this, EventArgs.Empty);
                break;

            case PowerModes.Resume:
                _settings.LastResumeUtc = DateTimeOffset.UtcNow;
                var diagnostics = CollectWakeDiagnostics();
                var analysis = AnalyzeWake(diagnostics);
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

    private void ArmResumeProtectionIfNeeded(WakeAnalysis analysis)
    {
        _resumeTimer.Stop();

        if (_settings.PolicyMode != PowerPolicyMode.FollowPowerPlan)
        {
            return;
        }

        if (!_settings.ResumeProtectionEnabled)
        {
            return;
        }

        if (_settings.ResumeProtectionOnlyForUnattendedWake && analysis.IsLikelyManualWake)
        {
            _logger.Info($"本次恢复被判定为疑似人工唤醒，已跳过自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
            return;
        }

        _resumeTimer.Interval = Math.Max(3, _settings.ResumeProtectionDelaySeconds) * 1000;
        _resumeTimer.Start();
        _logger.Warn($"已启动恢复保护，将在 {_settings.ResumeProtectionDelaySeconds} 秒后自动{DescribeResumeProtection(_settings.ResumeProtectionMode)}。");
    }

    private void ResumeTimerOnTick(object? sender, EventArgs e)
    {
        _resumeTimer.Stop();

        if (_settings.ResumeProtectionMode == ResumeProtectionMode.Hibernate)
        {
            HibernateNow();
        }
        else
        {
            SleepNow();
        }
    }

    private static string DescribeMode(PowerPolicyMode mode)
    {
        return mode == PowerPolicyMode.KeepAwakeIndefinitely ? "无限保持唤醒" : "遵循电源计划";
    }

    private static string DescribeResumeProtection(ResumeProtectionMode mode)
    {
        return mode == ResumeProtectionMode.Hibernate ? "休眠" : "睡眠";
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
            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            process.WaitForExit(3000);
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
        var text = diagnostics.ToLowerInvariant();

        if (ContainsAny(text, "power button", "电源按钮", "lid", "打开盖", "keyboard", "鼠标", "mouse"))
        {
            return new WakeAnalysis(true, "疑似人工唤醒（按键/开盖/鼠标键盘）");
        }

        if (ContainsAny(text, "wake timer", "timer", "task scheduler", "updateorchestrator", "maintenance activator"))
        {
            return new WakeAnalysis(false, "疑似软件或定时器唤醒");
        }

        if (ContainsAny(text, "device", "usb", "network adapter", "wake on lan", "pci", "网卡"))
        {
            return new WakeAnalysis(false, "疑似设备或网络唤醒");
        }

        if (ContainsAny(text, "history count - 0", "wake history count - 0", "no active wake timers"))
        {
            return new WakeAnalysis(false, "未识别到明确唤醒源，按非人工唤醒处理");
        }

        return new WakeAnalysis(false, "唤醒来源未知，按非人工唤醒处理");
    }

    private static bool ContainsAny(string source, params string[] values)
    {
        return values.Any(source.Contains);
    }

    private readonly record struct WakeAnalysis(bool IsLikelyManualWake, string Summary);
}
