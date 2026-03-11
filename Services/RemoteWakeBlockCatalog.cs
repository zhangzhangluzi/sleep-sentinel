namespace SleepSentinel.Services;

internal enum PowerRequestOverrideCallerType
{
    Process,
    Service
}

internal readonly record struct PowerRequestOverrideEntry(
    PowerRequestOverrideCallerType CallerType,
    string Name,
    string Product)
{
    public string CallerTypeArgument => CallerType == PowerRequestOverrideCallerType.Process ? "process" : "service";
}

internal static class RemoteWakeBlockCatalog
{
    public static IReadOnlyList<PowerRequestOverrideEntry> Entries { get; } =
    [
        new(PowerRequestOverrideCallerType.Process, "ToDesk.exe", "ToDesk"),
        new(PowerRequestOverrideCallerType.Process, "CoreToolsMgrHelper.exe", "ToDesk"),
        new(PowerRequestOverrideCallerType.Service, "ToDesk_Service", "ToDesk"),

        new(PowerRequestOverrideCallerType.Process, "SunloginClient.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "SunloginRemote.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "SunloginService.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "AweSun.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Process, "OldAweSun.exe", "向日葵 / AweSun"),
        new(PowerRequestOverrideCallerType.Service, "SunloginService", "向日葵 / AweSun"),

        new(PowerRequestOverrideCallerType.Process, "GameViewer.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerLauncher.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerServer.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerHealthd.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Process, "GameViewerService.exe", "GameViewer / UU远控"),
        new(PowerRequestOverrideCallerType.Service, "GameViewerService", "GameViewer / UU远控"),

        new(PowerRequestOverrideCallerType.Process, "AnyDesk.exe", "AnyDesk"),
        new(PowerRequestOverrideCallerType.Process, "ad_svc.exe", "AnyDesk"),
        new(PowerRequestOverrideCallerType.Service, "AnyDesk", "AnyDesk"),

        new(PowerRequestOverrideCallerType.Process, "TeamViewer.exe", "TeamViewer"),
        new(PowerRequestOverrideCallerType.Process, "TeamViewer_Service.exe", "TeamViewer"),
        new(PowerRequestOverrideCallerType.Service, "TeamViewer", "TeamViewer"),

        new(PowerRequestOverrideCallerType.Process, "RustDesk.exe", "RustDesk"),
        new(PowerRequestOverrideCallerType.Process, "rustdesk.exe", "RustDesk"),
        new(PowerRequestOverrideCallerType.Process, "RustDeskService.exe", "RustDesk"),
        new(PowerRequestOverrideCallerType.Service, "RustDesk", "RustDesk")
    ];

    public static string ProductSummary =>
        string.Join("、", Entries
            .Select(static entry => entry.Product)
            .Distinct(StringComparer.OrdinalIgnoreCase));
}
