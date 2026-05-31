using SleepSentinel.Services;
using SleepSentinel.UI;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading.Tasks;

namespace SleepSentinel;

internal static class Program
{
    private const string SingleInstanceMutexName = @"Global\SleepSentinel.SingleInstance";
    private const string ActivationEventName = @"Global\SleepSentinel.ActivateExisting";
    private const string TakeoverEventName = @"Global\SleepSentinel.TakeoverPrimary";
    private const string QuietStartupArg = "--quiet";
    private const string ElevatedTaskStartupArg = "--elevated-task";
    private const string ElevatedTaskStartupMarkerFileName = "elevated-task-start.marker";
    private const int TakeoverWaitMilliseconds = 20000;
    private const int ElevatedTaskBootstrapWaitMilliseconds = 8000;
    private const int TakeoverRetryIntervalMilliseconds = 250;
    private static readonly TimeSpan ElevatedTaskStartupMarkerMaxAge = TimeSpan.FromMinutes(2);

    [STAThread]
    private static void Main(string[] args)
    {
        var crashDirectory = GetCrashLogDirectory();
        var startupCrashLogPath = GetCrashLogPath();
        RegisterGlobalExceptionHandlers(crashDirectory);
        TrySetUnhandledExceptionMode(startupCrashLogPath);
        var startupMode = ResolveStartupMode(args);

        using var activationEvent = CreateGlobalEventWaitHandle(ActivationEventName);
        using var takeoverEvent = CreateGlobalEventWaitHandle(TakeoverEventName);
        using var singleInstance = AcquireSingleInstance(activationEvent, takeoverEvent, startupMode);
        if (singleInstance is null)
        {
            LogStartupTrace("检测到已有实例在运行，已尝试激活现有实例。");
            return;
        }

        try
        {
            ApplicationConfiguration.Initialize();

            var settingsStore = new SettingsStore();
            var logger = new FileLogger(settingsStore.BaseDirectory);
            var settings = settingsStore.Load();
            using var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application;
            using var controller = new PowerController(settingsStore, logger, settings);

            using var trayContext = new TrayApplicationContext(controller, logger, settingsStore, appIcon, activationEvent, takeoverEvent, startupMode.IsQuietStartup);
            Application.Run(trayContext);
        }
        catch (Exception ex)
        {
            var crashLogPath = GetCrashLogPath();
            LogUnhandledException(crashLogPath, ex);
            MessageBox.Show(
                $"SleepSentinel 启动失败：{ex.Message}，详见日志 {crashLogPath}",
                "SleepSentinel 启动失败",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static void LogStartupTrace(string message)
    {
        try
        {
            var crashPath = GetCrashLogPath();
            Directory.CreateDirectory(Path.GetDirectoryName(crashPath)!);
            File.AppendAllText(crashPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [Startup] {message}{Environment.NewLine}", Encoding.UTF8);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static string GetCrashLogDirectory()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SleepSentinel");
    }

    private static string GetCrashLogPath()
    {
        return Path.Combine(GetCrashLogDirectory(), "crash.log");
    }

    private static void TrySetUnhandledExceptionMode(string crashLogPath)
    {
        try
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        }
        catch (Exception ex)
        {
            LogUnhandledException(crashLogPath, ex);
        }
    }

    private static StartupMode ResolveStartupMode(string[] args)
    {
        var marker = IsRunningElevated()
            ? TryConsumeElevatedTaskStartupMarker()
            : default;
        var isElevatedTaskStartup = args.Any(static arg => string.Equals(arg, ElevatedTaskStartupArg, StringComparison.OrdinalIgnoreCase))
            || marker.Exists;
        var isQuietStartup = ShouldStartQuietly(args) && !marker.ShowMainWindow;
        return new StartupMode(isQuietStartup, isElevatedTaskStartup);
    }

    private static bool ShouldStartQuietly(string[] args)
    {
        return args.Any(static arg => string.Equals(arg, QuietStartupArg, StringComparison.OrdinalIgnoreCase));
    }

    private static ElevatedTaskStartupMarker TryConsumeElevatedTaskStartupMarker()
    {
        try
        {
            var markerPath = GetElevatedTaskStartupMarkerPath();
            if (!File.Exists(markerPath))
            {
                return default;
            }

            var markerAge = DateTimeOffset.Now - File.GetLastWriteTime(markerPath);
            var markerText = File.ReadAllText(markerPath);
            File.Delete(markerPath);
            if (markerAge < TimeSpan.Zero || markerAge > ElevatedTaskStartupMarkerMaxAge)
            {
                return default;
            }

            return new ElevatedTaskStartupMarker(
                Exists: true,
                ShowMainWindow: markerText.Contains("mode=show", StringComparison.OrdinalIgnoreCase));
        }
        catch (IOException)
        {
            return default;
        }
        catch (UnauthorizedAccessException)
        {
            return default;
        }
    }

    private static string GetElevatedTaskStartupMarkerPath()
    {
        return Path.Combine(GetCrashLogDirectory(), ElevatedTaskStartupMarkerFileName);
    }

    private static void RegisterGlobalExceptionHandlers(string logDirectory)
    {
        var appFolder = string.IsNullOrWhiteSpace(logDirectory)
            ? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            : logDirectory;
        var crashPath = Path.Combine(appFolder, "crash.log");

        Application.ThreadException += (_, e) =>
        {
            LogUnhandledException(crashPath, e.Exception);
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            LogUnhandledException(crashPath, e.Exception);
            e.SetObserved();
        };

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogUnhandledException(crashPath, ex);
            }
            else
            {
                LogUnhandledException(crashPath, new Exception($"Unhandled exception: {e.ExceptionObject}"));
            }
        };
    }

    private static void LogUnhandledException(string crashLogPath, Exception ex)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(crashLogPath)!);
            var message = new StringBuilder();
            message.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [Unhandled] {ex}");
            File.AppendAllText(crashLogPath, message.ToString(), Encoding.UTF8);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static Mutex? AcquireSingleInstance(EventWaitHandle activationEvent, EventWaitHandle takeoverEvent, StartupMode startupMode)
    {
        var singleInstance = TryCreatePrimaryInstanceMutex();
        if (singleInstance is not null)
        {
            singleInstance = ResolveUnexpectedDuplicateProcess(singleInstance, activationEvent, takeoverEvent);
            if (singleInstance is null)
            {
                return null;
            }

            singleInstance = TrySwitchPrimaryRoleToElevatedTask(singleInstance, startupMode);
            if (singleInstance is null)
            {
                return null;
            }

            return singleInstance;
        }

        if (IsRunningElevated() && TryRequestExistingInstanceTakeover(takeoverEvent))
        {
            singleInstance = WaitForPrimaryInstanceRelease();
            if (singleInstance is not null)
            {
                return singleInstance;
            }

            LogStartupTrace("已请求低权限实例退出，但等待主实例锁释放超时。");
        }

        TryActivateExistingInstance(activationEvent);
        return null;
    }

    private static Mutex? TrySwitchPrimaryRoleToElevatedTask(Mutex singleInstance, StartupMode startupMode)
    {
        if (IsRunningElevated() || startupMode.IsElevatedTaskStartup || !startupMode.IsQuietStartup)
        {
            return singleInstance;
        }

        if (!AutostartManager.TryStartElevatedScheduledTaskForCurrentExecutable(showMainWindow: !startupMode.IsQuietStartup, out _))
        {
            return singleInstance;
        }

        singleInstance.Dispose();
        return WaitForPrimaryInstanceToAppear()
            ? null
            : TryCreatePrimaryInstanceMutex();
    }

    private static Mutex? TryCreatePrimaryInstanceMutex()
    {
        try
        {
            var singleInstance = MutexAcl.Create(true, SingleInstanceMutexName, out var createdNew, CreateGlobalMutexSecurity());
            if (createdNew)
            {
                return singleInstance;
            }

            singleInstance.Dispose();
        }
        catch (UnauthorizedAccessException)
        {
        }

        return null;
    }

    private static Mutex? ResolveUnexpectedDuplicateProcess(Mutex singleInstance, EventWaitHandle activationEvent, EventWaitHandle takeoverEvent)
    {
        if (!IsAnotherSleepSentinelProcessRunning())
        {
            return singleInstance;
        }

        if (IsRunningElevated() && TryRequestExistingInstanceTakeover(takeoverEvent) && WaitForOtherSleepSentinelProcessesToExit())
        {
            LogStartupTrace("检测到已有低权限进程，已接管为当前高权限主实例。");
            return singleInstance;
        }

        LogStartupTrace("检测到已有 SleepSentinel 进程，但当前实例仍获得了主实例锁；已放弃当前实例以避免重复窗口。");
        TryActivateExistingInstance(activationEvent);
        singleInstance.Dispose();
        return null;
    }

    private static Mutex? WaitForPrimaryInstanceRelease()
    {
        var deadlineTick = Environment.TickCount64 + TakeoverWaitMilliseconds;
        while (Environment.TickCount64 < deadlineTick)
        {
            Thread.Sleep(TakeoverRetryIntervalMilliseconds);
            var singleInstance = TryCreatePrimaryInstanceMutex();
            if (singleInstance is not null)
            {
                return singleInstance;
            }
        }

        return null;
    }

    private static bool WaitForPrimaryInstanceToAppear()
    {
        var deadlineTick = Environment.TickCount64 + ElevatedTaskBootstrapWaitMilliseconds;
        while (Environment.TickCount64 < deadlineTick)
        {
            Thread.Sleep(TakeoverRetryIntervalMilliseconds);
            if (TryOpenPrimaryInstanceMutex())
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryOpenPrimaryInstanceMutex()
    {
        try
        {
            if (MutexAcl.TryOpenExisting(SingleInstanceMutexName, MutexRights.Synchronize, out var singleInstance))
            {
                singleInstance.Dispose();
                return true;
            }
        }
        catch (WaitHandleCannotBeOpenedException)
        {
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }

        return false;
    }

    private static bool WaitForOtherSleepSentinelProcessesToExit()
    {
        var deadlineTick = Environment.TickCount64 + TakeoverWaitMilliseconds;
        while (Environment.TickCount64 < deadlineTick)
        {
            if (!IsAnotherSleepSentinelProcessRunning())
            {
                return true;
            }

            Thread.Sleep(TakeoverRetryIntervalMilliseconds);
        }

        return !IsAnotherSleepSentinelProcessRunning();
    }

    private static bool IsAnotherSleepSentinelProcessRunning()
    {
        var currentProcessId = Environment.ProcessId;
        var currentPath = TryGetCurrentProcessPath();
        var processName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);

        foreach (var process in Process.GetProcessesByName(processName))
        {
            using (process)
            {
                if (process.Id == currentProcessId)
                {
                    continue;
                }

                try
                {
                    if (process.HasExited)
                    {
                        continue;
                    }
                }
                catch (InvalidOperationException)
                {
                    continue;
                }

                var processPath = TryGetProcessPath(process);
                if (string.IsNullOrWhiteSpace(currentPath)
                    || string.IsNullOrWhiteSpace(processPath)
                    || string.Equals(currentPath, processPath, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static string? TryGetCurrentProcessPath()
    {
        try
        {
            return Environment.ProcessPath;
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetProcessPath(Process process)
    {
        try
        {
            return process.MainModule?.FileName;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return null;
        }
        catch (NotSupportedException)
        {
            return null;
        }
    }

    private static EventWaitHandle CreateGlobalEventWaitHandle(string name)
    {
        try
        {
            return EventWaitHandleAcl.Create(
                false,
                EventResetMode.AutoReset,
                name,
                out _,
                CreateGlobalEventSecurity());
        }
        catch (UnauthorizedAccessException)
        {
            return EventWaitHandleAcl.OpenExisting(
                name,
                EventWaitHandleRights.Modify | EventWaitHandleRights.Synchronize);
        }
    }

    private static MutexSecurity CreateGlobalMutexSecurity()
    {
        var security = new MutexSecurity();
        security.AddAccessRule(new MutexAccessRule(
            CreateCrossIntegrityAccessSid(),
            MutexRights.FullControl,
            AccessControlType.Allow));
        return security;
    }

    private static EventWaitHandleSecurity CreateGlobalEventSecurity()
    {
        var security = new EventWaitHandleSecurity();
        security.AddAccessRule(new EventWaitHandleAccessRule(
            CreateCrossIntegrityAccessSid(),
            EventWaitHandleRights.FullControl,
            AccessControlType.Allow));
        return security;
    }

    private static SecurityIdentifier CreateCrossIntegrityAccessSid()
    {
        return new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null);
    }

    private static void TryActivateExistingInstance(EventWaitHandle activationEvent)
    {
        try
        {
            activationEvent.Set();
        }
        catch (UnauthorizedAccessException)
        {
            LogStartupTrace("检测到已有实例，但当前实例没有权限发送唤醒信号。");
        }
    }

    private static bool TryRequestExistingInstanceTakeover(EventWaitHandle takeoverEvent)
    {
        try
        {
            takeoverEvent.Set();
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static bool IsRunningElevated()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private readonly record struct StartupMode(bool IsQuietStartup, bool IsElevatedTaskStartup);

    private readonly record struct ElevatedTaskStartupMarker(bool Exists, bool ShowMainWindow);
}
