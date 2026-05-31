using SleepSentinel.Services;
using SleepSentinel.UI;
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
        var isQuietStartup = ShouldStartQuietly(args);
        var isElevatedTaskStartup = IsElevatedTaskStartup(args);

        using var activationEvent = CreateGlobalEventWaitHandle(ActivationEventName);
        using var takeoverEvent = CreateGlobalEventWaitHandle(TakeoverEventName);
        using var singleInstance = AcquireSingleInstance(activationEvent, takeoverEvent, isElevatedTaskStartup);
        if (singleInstance is null)
        {
            LogStartupTrace("检测到已有实例在运行，已尝试激活现有实例。");
            if (!isQuietStartup)
            {
                ShowExistingInstanceHint();
            }
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

            using var trayContext = new TrayApplicationContext(controller, logger, settingsStore, appIcon, activationEvent, takeoverEvent);
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

    private static bool ShouldStartQuietly(string[] args)
    {
        return args.Any(static arg => string.Equals(arg, QuietStartupArg, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsElevatedTaskStartup(string[] args)
    {
        return args.Any(static arg => string.Equals(arg, ElevatedTaskStartupArg, StringComparison.OrdinalIgnoreCase))
            || IsRunningElevated() && TryConsumeElevatedTaskStartupMarker();
    }

    private static bool TryConsumeElevatedTaskStartupMarker()
    {
        try
        {
            var markerPath = GetElevatedTaskStartupMarkerPath();
            if (!File.Exists(markerPath))
            {
                return false;
            }

            var markerAge = DateTimeOffset.Now - File.GetLastWriteTime(markerPath);
            File.Delete(markerPath);
            return markerAge >= TimeSpan.Zero && markerAge <= ElevatedTaskStartupMarkerMaxAge;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static string GetElevatedTaskStartupMarkerPath()
    {
        return Path.Combine(GetCrashLogDirectory(), ElevatedTaskStartupMarkerFileName);
    }

    private static void ShowExistingInstanceHint()
    {
        try
        {
            MessageBox.Show(
                "SleepSentinel 已在后台运行中，当前点击属于唤起已有实例。",
                "SleepSentinel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch
        {
            // Ignore UI errors when session or desktop is not ready.
        }
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

    private static Mutex? AcquireSingleInstance(EventWaitHandle activationEvent, EventWaitHandle takeoverEvent, bool isElevatedTaskStartup)
    {
        var singleInstance = TryCreatePrimaryInstanceMutex();
        if (singleInstance is not null)
        {
            singleInstance = TrySwitchPrimaryRoleToElevatedTask(singleInstance, isElevatedTaskStartup);
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

    private static Mutex? TrySwitchPrimaryRoleToElevatedTask(Mutex singleInstance, bool isElevatedTaskStartup)
    {
        if (IsRunningElevated() || isElevatedTaskStartup)
        {
            return singleInstance;
        }

        if (!AutostartManager.TryStartElevatedScheduledTaskForCurrentExecutable(out _))
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
            MessageBox.Show(
                "SleepSentinel 已经在运行，但当前实例无法唤回它。",
                "SleepSentinel",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
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
}
