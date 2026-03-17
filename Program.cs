using SleepSentinel.Services;
using SleepSentinel.UI;

namespace SleepSentinel;

internal static class Program
{
    private const string SingleInstanceMutexName = @"Global\SleepSentinel.SingleInstance";
    private const string ActivationEventName = @"Global\SleepSentinel.ActivateExisting";

    [STAThread]
    private static void Main()
    {
        using var activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ActivationEventName);
        using var singleInstance = new Mutex(true, SingleInstanceMutexName, out var createdNew);
        if (!createdNew)
        {
            TryActivateExistingInstance(activationEvent);
            return;
        }

        ApplicationConfiguration.Initialize();

        var settingsStore = new SettingsStore();
        var logger = new FileLogger(settingsStore.BaseDirectory);
        var settings = settingsStore.Load();
        using var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application;
        using var controller = new PowerController(settingsStore, logger, settings);
        using var trayContext = new TrayApplicationContext(controller, logger, settingsStore, appIcon, activationEvent);

        Application.Run(trayContext);
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
}
