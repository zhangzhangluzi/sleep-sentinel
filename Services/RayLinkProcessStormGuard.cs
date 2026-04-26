using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Net.NetworkInformation;
using System.ServiceProcess;

namespace SleepSentinel.Services;

internal sealed class RayLinkProcessStormGuard : IDisposable
{
    private const int WatchProcessStormThreshold = 5;
    private const int TotalProcessStormThreshold = 25;
    private const int PortConnectionStormThreshold = 30;
    private const int ServiceCrashStormThreshold = 3;
    private const int ServiceCrashCorroboratingProcessThreshold = 15;
    private const int ServiceCrashCorroboratingPortThreshold = 20;
    private static readonly TimeSpan ScanPeriod = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan InitialScanDelay = TimeSpan.FromSeconds(8);
    private static readonly TimeSpan BurstScanPeriod = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan BurstScanDuration = TimeSpan.FromSeconds(90);
    private static readonly TimeSpan SleepIsolationScanPeriod = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan SleepIsolationRestoreDelay = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan ContainmentCooldown = TimeSpan.FromSeconds(20);
    private static readonly TimeSpan ServiceStopTimeout = TimeSpan.FromSeconds(8);
    private static readonly TimeSpan ServiceStartTimeout = TimeSpan.FromSeconds(8);
    private static readonly string[] RayLinkProcessNames =
    [
        "RayLink",
        "RayLinkService",
        "RayLinkWatch",
        "RayLinkCapturer"
    ];

    private readonly FileLogger _logger;
    private readonly System.Threading.Timer _timer;
    private readonly object _sync = new();
    private DateTimeOffset _lastContainmentUtc = DateTimeOffset.MinValue;
    private DateTimeOffset _burstScanUntilUtc = DateTimeOffset.MinValue;
    private int _restoreGeneration;
    private int _scanInProgress;
    private bool _enabled;
    private bool _autoContain;
    private bool _isolateDuringSleep;
    private bool _sleepIsolationActive;
    private bool _restoreRayLinkServiceAfterSleep;
    private bool _disposed;
    private string _currentSummary = "RayLink 进程风暴守护未启用";
    private string _currentQuickState = "当前：未启用";

    public RayLinkProcessStormGuard(FileLogger logger)
    {
        _logger = logger;
        _timer = new System.Threading.Timer(
            static state => ((RayLinkProcessStormGuard)state!).OnTimerElapsed(),
            this,
            Timeout.InfiniteTimeSpan,
            Timeout.InfiniteTimeSpan);
    }

    public event EventHandler? StateChanged;

    public event EventHandler<string>? NotificationRequested;

    public string CurrentSummary
    {
        get
        {
            lock (_sync)
            {
                return _currentSummary;
            }
        }
    }

    public string CurrentQuickState
    {
        get
        {
            lock (_sync)
            {
                return _currentQuickState;
            }
        }
    }

    public void UpdateSettings(bool enabled, bool autoContain, bool isolateDuringSleep)
    {
        lock (_sync)
        {
            if (_disposed)
            {
                return;
            }

            _enabled = enabled;
            _autoContain = autoContain;
            _isolateDuringSleep = isolateDuringSleep;
            if (!isolateDuringSleep)
            {
                _sleepIsolationActive = false;
                _restoreRayLinkServiceAfterSleep = false;
                _restoreGeneration++;
            }

            if (!enabled)
            {
                _burstScanUntilUtc = DateTimeOffset.MinValue;
                _sleepIsolationActive = false;
                _restoreRayLinkServiceAfterSleep = false;
                _restoreGeneration++;
                _timer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
                SetStateUnsafe("RayLink 进程风暴守护未启用", "当前：未启用");
                return;
            }

            SetStateUnsafe(
                autoContain
                    ? "RayLink 进程风暴守护已启用；仅在异常风暴时停止服务并结束 RayLink 进程树，不会禁用服务"
                    : "RayLink 进程风暴守护已启用；当前仅监控，不自动止血",
                "当前：监控中");
            _timer.Change(InitialScanDelay, ScanPeriod);
        }
    }

    public void BeginSleepIsolation(string reason)
    {
        bool enabled;
        bool isolateDuringSleep;
        lock (_sync)
        {
            enabled = _enabled;
            isolateDuringSleep = _isolateDuringSleep;
        }

        if (!enabled || !isolateDuringSleep)
        {
            TriggerHighFrequencyScan(reason);
            return;
        }

        var shouldRestoreService = IsRayLinkServiceRunning();
        lock (_sync)
        {
            if (_disposed || !_enabled || !_isolateDuringSleep)
            {
                return;
            }

            _sleepIsolationActive = true;
            _restoreRayLinkServiceAfterSleep |= shouldRestoreService;
            _restoreGeneration++;
            SetStateUnsafe(
                $"RayLink 睡眠隔离已启用；触发：{reason}；睡眠期间会暂停并持续压制 RayLink，不禁用服务",
                "当前：睡眠隔离中");
            _timer.Change(TimeSpan.Zero, SleepIsolationScanPeriod);
        }

        _logger.Warn($"RayLink 睡眠隔离启动：{reason}。将暂停 RayLink，避免睡眠期间被它反复拉起。");
        EnforceSleepIsolationIfNeeded("睡眠隔离启动");
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void ScheduleRestoreAfterManualResume(string reason)
    {
        int generation;
        var shouldSchedule = false;
        lock (_sync)
        {
            if (_disposed || !_enabled || !_sleepIsolationActive)
            {
                return;
            }

            generation = ++_restoreGeneration;
            shouldSchedule = true;
            SetStateUnsafe(
                $"RayLink 睡眠隔离等待恢复；触发：{reason}；将在 {SleepIsolationRestoreDelay.TotalSeconds:F0} 秒后恢复服务",
                "当前：等待恢复");
            _timer.Change(TimeSpan.Zero, SleepIsolationScanPeriod);
        }

        if (!shouldSchedule)
        {
            return;
        }

        _logger.Info($"检测到人工恢复（{reason}），RayLink 将在 {SleepIsolationRestoreDelay.TotalSeconds:F0} 秒后恢复。");
        StateChanged?.Invoke(this, EventArgs.Empty);
        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(SleepIsolationRestoreDelay).ConfigureAwait(false);
                CompleteSleepIsolationRestore(generation);
            }
            catch (Exception ex)
            {
                _logger.Warn($"RayLink 睡眠隔离恢复任务失败：{ex.Message}");
            }
        });
    }

    public void TriggerHighFrequencyScan(string reason)
    {
        var shouldLog = false;
        lock (_sync)
        {
            if (_disposed || !_enabled)
            {
                return;
            }

            var previousBurstUntilUtc = _burstScanUntilUtc;
            var nextBurstUntilUtc = DateTimeOffset.UtcNow.Add(BurstScanDuration);
            if (nextBurstUntilUtc > _burstScanUntilUtc)
            {
                _burstScanUntilUtc = nextBurstUntilUtc;
            }

            shouldLog = previousBurstUntilUtc <= DateTimeOffset.UtcNow;
            SetStateUnsafe(
                _autoContain
                    ? $"RayLink 进程风暴守护已进入高频扫描；触发：{reason}；异常时自动止血，不禁用服务"
                    : $"RayLink 进程风暴守护已进入高频扫描；触发：{reason}；当前仅监控",
                "当前：高频扫描");
            _timer.Change(TimeSpan.Zero, BurstScanPeriod);
        }

        if (shouldLog)
        {
            _logger.Info($"RayLink 进程风暴守护进入高频扫描：{reason}。");
        }

        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void Dispose()
    {
        lock (_sync)
        {
            _disposed = true;
            _timer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        }

        _timer.Dispose();
    }

    private void OnTimerElapsed()
    {
        if (Interlocked.Exchange(ref _scanInProgress, 1) != 0)
        {
            return;
        }

        try
        {
            bool enabled;
            bool autoContain;
            lock (_sync)
            {
                if (_disposed)
                {
                    return;
                }

                enabled = _enabled;
                autoContain = _autoContain;
            }

            if (!enabled)
            {
                return;
            }

            if (IsSleepIsolationActive())
            {
                EnforceSleepIsolationIfNeeded("睡眠隔离持续拦截");
                return;
            }

            var snapshot = CollectSnapshot();
            if (!snapshot.IsStorm)
            {
                SetState(snapshot.BuildHealthySummary(autoContain), snapshot.BuildHealthyQuickState());
                return;
            }

            SetState(snapshot.BuildStormSummary(autoContain), "当前：检测到风暴");
            if (!autoContain)
            {
                _logger.Warn(snapshot.BuildStormSummary(autoContain));
                NotificationRequested?.Invoke(this, snapshot.BuildStormNotification(autoContained: false));
                return;
            }

            if (DateTimeOffset.UtcNow - _lastContainmentUtc < ContainmentCooldown)
            {
                return;
            }

            _lastContainmentUtc = DateTimeOffset.UtcNow;
            ContainStorm(snapshot);
        }
        catch (Exception ex)
        {
            _logger.Warn($"RayLink 进程风暴守护扫描失败：{ex.Message}");
            SetState("RayLink 进程风暴守护扫描失败，请查看日志", "当前：扫描失败");
        }
        finally
        {
            Interlocked.Exchange(ref _scanInProgress, 0);
            RescheduleIfBurstExpired();
        }
    }

    private void RescheduleIfBurstExpired()
    {
        lock (_sync)
        {
            if (_disposed || !_enabled)
            {
                return;
            }

            if (!_sleepIsolationActive
                && _burstScanUntilUtc != DateTimeOffset.MinValue
                && DateTimeOffset.UtcNow >= _burstScanUntilUtc)
            {
                _burstScanUntilUtc = DateTimeOffset.MinValue;
                _timer.Change(ScanPeriod, ScanPeriod);
                _logger.Info("RayLink 进程风暴守护已退出高频扫描，恢复常规轮询。");
            }
        }
    }

    private bool IsSleepIsolationActive()
    {
        lock (_sync)
        {
            return !_disposed && _enabled && _sleepIsolationActive;
        }
    }

    private void EnforceSleepIsolationIfNeeded(string reason)
    {
        if (!HasRayLinkActivity())
        {
            SetState("RayLink 睡眠隔离中：RayLink 已暂停，睡眠期间不会放行它重新拉起", "当前：睡眠隔离中");
            return;
        }

        var serviceResult = StopRayLinkService();
        var processResult = KillRayLinkProcessTree();
        var summary =
            $"RayLink 睡眠隔离拦截：{reason}；{serviceResult}，结束 RayLink 进程 {processResult.KilledCount} 个；不禁用服务，人工恢复后再延迟恢复";

        if (processResult.Failures.Count > 0)
        {
            summary += $"；失败 {processResult.Failures.Count} 项";
            _logger.Warn($"RayLink 睡眠隔离时部分进程结束失败：{Environment.NewLine}{string.Join(Environment.NewLine, processResult.Failures)}");
        }

        _logger.Warn(summary);
        SetState(summary, "当前：已隔离");
    }

    private void CompleteSleepIsolationRestore(int generation)
    {
        var shouldStartService = false;
        lock (_sync)
        {
            if (_disposed || !_enabled || !_sleepIsolationActive || generation != _restoreGeneration)
            {
                return;
            }

            _sleepIsolationActive = false;
            shouldStartService = _restoreRayLinkServiceAfterSleep;
            _restoreRayLinkServiceAfterSleep = false;
            _timer.Change(ScanPeriod, ScanPeriod);
        }

        var restoreResult = shouldStartService
            ? StartRayLinkService()
            : "睡前 RayLinkService 未运行，本次不自动启动";
        _logger.Info($"RayLink 睡眠隔离已结束：{restoreResult}。");
        SetState($"RayLink 睡眠隔离已结束：{restoreResult}", "当前：监控中");
    }

    private RayLinkStormSnapshot CollectSnapshot()
    {
        var processes = EnumerateProcessesSafely().ToArray();
        try
        {
            var watchCount = processes.Count(static process => IsProcessNamed(process, "RayLinkWatch"));
            var totalRayLinkCount = processes.Count(IsRayLinkManagedProcess);
            var portConnectionCount = CountLoopbackPortConnections(6511);
            var serviceCrashCount = CountRecentServiceCrashes();
            var hasCrashCorroboratingActivity = watchCount > 1
                || totalRayLinkCount >= ServiceCrashCorroboratingProcessThreshold
                || portConnectionCount >= ServiceCrashCorroboratingPortThreshold;
            var isStorm = watchCount >= WatchProcessStormThreshold
                || totalRayLinkCount >= TotalProcessStormThreshold
                || portConnectionCount >= PortConnectionStormThreshold
                || (serviceCrashCount >= ServiceCrashStormThreshold && hasCrashCorroboratingActivity);

            return new RayLinkStormSnapshot(watchCount, totalRayLinkCount, portConnectionCount, serviceCrashCount, isStorm);
        }
        finally
        {
            foreach (var process in processes)
            {
                process.Dispose();
            }
        }
    }

    private void ContainStorm(RayLinkStormSnapshot snapshot)
    {
        var serviceResult = StopRayLinkService();
        var processResult = KillRayLinkProcessTree();
        var summary =
            $"{snapshot.BuildStormSummary(autoContain: true)}；已止血：{serviceResult}，结束 RayLink 进程 {processResult.KilledCount} 个；未禁用 RayLinkService，可手动重新打开 RayLink";

        if (processResult.Failures.Count > 0)
        {
            summary += $"；失败 {processResult.Failures.Count} 项";
            _logger.Warn($"RayLink 进程风暴止血时部分进程结束失败：{Environment.NewLine}{string.Join(Environment.NewLine, processResult.Failures)}");
        }

        _logger.Warn(summary);
        SetState(summary, "当前：已止血");
        NotificationRequested?.Invoke(this, snapshot.BuildStormNotification(autoContained: true));
    }

    private static KillResult KillRayLinkProcessTree()
    {
        var killedCount = 0;
        var failures = new List<string>();
        var processes = EnumerateProcessesSafely().ToArray();
        foreach (var process in processes)
        {
            using (process)
            {
                if (!IsRayLinkManagedProcess(process))
                {
                    continue;
                }

                try
                {
                    process.Kill(entireProcessTree: true);
                    killedCount++;
                }
                catch (InvalidOperationException)
                {
                }
                catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or NotSupportedException)
                {
                    failures.Add($"{process.ProcessName}({SafeProcessId(process)}): {ex.Message}");
                }
            }
        }

        return new KillResult(killedCount, failures);
    }

    private static string StopRayLinkService()
    {
        try
        {
            using var service = new ServiceController("RayLinkService");
            if (service.Status is ServiceControllerStatus.Stopped or ServiceControllerStatus.StopPending)
            {
                return "RayLinkService 已停止";
            }

            service.Stop();
            service.WaitForStatus(ServiceControllerStatus.Stopped, ServiceStopTimeout);
            return service.Status == ServiceControllerStatus.Stopped
                ? "RayLinkService 已停止"
                : $"RayLinkService 停止超时（当前 {service.Status}）";
        }
        catch (InvalidOperationException)
        {
            return "未检测到 RayLinkService";
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or System.ServiceProcess.TimeoutException or System.TimeoutException)
        {
            return $"RayLinkService 停止失败：{ex.Message}";
        }
    }

    private static string StartRayLinkService()
    {
        try
        {
            using var service = new ServiceController("RayLinkService");
            if (service.Status is ServiceControllerStatus.Running or ServiceControllerStatus.StartPending)
            {
                return "RayLinkService 已运行";
            }

            service.Start();
            service.WaitForStatus(ServiceControllerStatus.Running, ServiceStartTimeout);
            return service.Status == ServiceControllerStatus.Running
                ? "RayLinkService 已恢复运行"
                : $"RayLinkService 启动超时（当前 {service.Status}）";
        }
        catch (InvalidOperationException)
        {
            return "未检测到 RayLinkService";
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or System.ServiceProcess.TimeoutException or System.TimeoutException)
        {
            return $"RayLinkService 启动失败：{ex.Message}";
        }
    }

    private static bool IsRayLinkServiceRunning()
    {
        try
        {
            using var service = new ServiceController("RayLinkService");
            return service.Status is ServiceControllerStatus.Running or ServiceControllerStatus.StartPending;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return false;
        }
    }

    private static bool HasRayLinkActivity()
    {
        if (IsRayLinkServiceRunning())
        {
            return true;
        }

        var processes = EnumerateProcessesSafely().ToArray();
        try
        {
            return processes.Any(IsRayLinkManagedProcess) || CountLoopbackPortConnections(6511) > 0;
        }
        finally
        {
            foreach (var process in processes)
            {
                process.Dispose();
            }
        }
    }

    private void SetState(string summary, string quickState)
    {
        var changed = false;
        lock (_sync)
        {
            changed = SetStateUnsafe(summary, quickState);
        }

        if (changed)
        {
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private bool SetStateUnsafe(string summary, string quickState)
    {
        if (string.Equals(_currentSummary, summary, StringComparison.Ordinal)
            && string.Equals(_currentQuickState, quickState, StringComparison.Ordinal))
        {
            return false;
        }

        _currentSummary = summary;
        _currentQuickState = quickState;
        return true;
    }

    private static IEnumerable<Process> EnumerateProcessesSafely()
    {
        try
        {
            return Process.GetProcesses();
        }
        catch
        {
            return [];
        }
    }

    private static bool IsRayLinkManagedProcess(Process process)
    {
        if (RayLinkProcessNames.Any(name => IsProcessNamed(process, name)))
        {
            return true;
        }

        return IsProcessNamed(process, "speedtest") && IsUnderRayLinkDirectory(process);
    }

    private static bool IsProcessNamed(Process process, string name)
    {
        try
        {
            return process.ProcessName.Equals(name, StringComparison.OrdinalIgnoreCase);
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    private static bool IsUnderRayLinkDirectory(Process process)
    {
        try
        {
            var fileName = process.MainModule?.FileName;
            return !string.IsNullOrWhiteSpace(fileName)
                && fileName.Contains(@"\RayLink\", StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception or NotSupportedException)
        {
            return false;
        }
    }

    private static int CountLoopbackPortConnections(int port)
    {
        try
        {
            return IPGlobalProperties.GetIPGlobalProperties()
                .GetActiveTcpConnections()
                .Count(connection =>
                    connection.LocalEndPoint.Port == port
                    && IPAddressIsLoopback(connection.LocalEndPoint.Address));
        }
        catch
        {
            return 0;
        }
    }

    private static bool IPAddressIsLoopback(System.Net.IPAddress address)
    {
        return System.Net.IPAddress.IsLoopback(address);
    }

    private static int CountRecentServiceCrashes()
    {
        var sinceUtc = DateTimeOffset.UtcNow.AddMinutes(-10);
        var count = 0;
        try
        {
            var query = new EventLogQuery("System", PathType.LogName)
            {
                ReverseDirection = true
            };

            using var reader = new EventLogReader(query);
            for (EventRecord? record = reader.ReadEvent(); record is not null; record = reader.ReadEvent())
            {
                using (record)
                {
                    if (record.TimeCreated is not { } createdAt)
                    {
                        continue;
                    }

                    if (new DateTimeOffset(createdAt) < sinceUtc)
                    {
                        break;
                    }

                    if (record.Id != 7034 || !string.Equals(record.ProviderName, "Service Control Manager", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var message = record.FormatDescription();
                    if (message?.Contains("RayLink Service", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        count++;
                    }
                }
            }
        }
        catch
        {
            return 0;
        }

        return count;
    }

    private static int SafeProcessId(Process process)
    {
        try
        {
            return process.Id;
        }
        catch (InvalidOperationException)
        {
            return 0;
        }
    }

    private readonly record struct KillResult(int KilledCount, IReadOnlyList<string> Failures);

    private readonly record struct RayLinkStormSnapshot(
        int WatchCount,
        int TotalRayLinkCount,
        int PortConnectionCount,
        int ServiceCrashCount,
        bool IsStorm)
    {
        public string BuildHealthySummary(bool autoContain)
        {
            return autoContain
                ? $"RayLink 进程风暴守护正常：Watch={WatchCount}，RayLink总数={TotalRayLinkCount}，6511连接={PortConnectionCount}，近10分钟服务崩溃={ServiceCrashCount}；异常时自动止血，不禁用服务"
                : $"RayLink 进程风暴守护正常：Watch={WatchCount}，RayLink总数={TotalRayLinkCount}，6511连接={PortConnectionCount}，近10分钟服务崩溃={ServiceCrashCount}；当前仅监控";
        }

        public string BuildHealthyQuickState()
        {
            return TotalRayLinkCount == 0 ? "当前：未运行" : $"当前：正常 ({TotalRayLinkCount})";
        }

        public string BuildStormSummary(bool autoContain)
        {
            var action = autoContain ? "将自动止血" : "仅记录不止血";
            return $"检测到 RayLink 进程风暴：Watch={WatchCount}，RayLink总数={TotalRayLinkCount}，6511连接={PortConnectionCount}，近10分钟服务崩溃={ServiceCrashCount}，{action}";
        }

        public string BuildStormNotification(bool autoContained)
        {
            return autoContained
                ? $"RayLink 进程风暴已止血：Watch={WatchCount}，总数={TotalRayLinkCount}。RayLink 已暂停但未禁用，可手动重新打开。"
                : $"检测到 RayLink 进程风暴：Watch={WatchCount}，总数={TotalRayLinkCount}。当前仅监控，未结束进程。";
        }
    }
}
