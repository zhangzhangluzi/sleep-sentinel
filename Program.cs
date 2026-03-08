using SleepSentinel.Services;
using SleepSentinel.UI;

namespace SleepSentinel;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        using var singleInstance = new Mutex(true, @"Global\SleepSentinel.SingleInstance", out var createdNew);
        if (!createdNew)
        {
            MessageBox.Show("SleepSentinel 已经在运行。", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ApplicationConfiguration.Initialize();

        var settingsStore = new SettingsStore();
        var logger = new FileLogger(settingsStore.BaseDirectory);
        var settings = settingsStore.Load();
        using var controller = new PowerController(settingsStore, logger, settings);
        using var trayContext = new TrayApplicationContext(controller, logger, settingsStore);

        Application.Run(trayContext);
    }
}
