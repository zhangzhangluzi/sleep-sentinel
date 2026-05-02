using System.Reflection;

namespace SleepSentinel.Services;

internal static class AppVersionInfo
{
    public static string Version { get; } = ResolveVersion();

    public static string FileVersion { get; } = ResolveFileVersion();

    public static string ShortDisplayVersion => $"v{Version}";

    public static string FullDisplayVersion => $"SleepSentinel {ShortDisplayVersion}";

    public static string DetailedDisplayVersion => FileVersion.Equals(Version, StringComparison.OrdinalIgnoreCase)
        ? ShortDisplayVersion
        : $"{ShortDisplayVersion}（文件版本 {FileVersion}）";

    public static string WindowTitle => $"{FullDisplayVersion}";

    public static string TrayTitle => $"SleepSentinel {ShortDisplayVersion}";

    private static string ResolveVersion()
    {
        var assembly = typeof(AppVersionInfo).Assembly;
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;
        var version = StripBuildMetadata(informationalVersion);
        if (!string.IsNullOrWhiteSpace(version))
        {
            return version;
        }

        version = StripBuildMetadata(assembly.GetName().Version?.ToString());
        return string.IsNullOrWhiteSpace(version) ? "unknown" : version;
    }

    private static string ResolveFileVersion()
    {
        var assembly = typeof(AppVersionInfo).Assembly;
        var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
        fileVersion = StripBuildMetadata(fileVersion);
        return string.IsNullOrWhiteSpace(fileVersion) ? Version : fileVersion;
    }

    private static string StripBuildMetadata(string? version)
    {
        if (string.IsNullOrWhiteSpace(version))
        {
            return string.Empty;
        }

        var metadataIndex = version.IndexOf('+');
        return metadataIndex < 0 ? version.Trim() : version[..metadataIndex].Trim();
    }
}
