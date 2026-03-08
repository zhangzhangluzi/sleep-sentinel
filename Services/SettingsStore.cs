using System.Text.Json;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class SettingsStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true
    };
    private readonly object _sync = new();

    public string BaseDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SleepSentinel");

    public string SettingsPath => Path.Combine(BaseDirectory, "settings.json");

    public AppSettings Load()
    {
        lock (_sync)
        {
            Directory.CreateDirectory(BaseDirectory);

            if (!File.Exists(SettingsPath))
            {
                return CreateDefaultSettingsUnsafe();
            }

            try
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json, SerializerOptions) ?? BackupAndResetSettingsUnsafe();
            }
            catch (JsonException)
            {
                return BackupAndResetSettingsUnsafe();
            }
            catch (NotSupportedException)
            {
                return BackupAndResetSettingsUnsafe();
            }
            catch (IOException)
            {
                return new AppSettings();
            }
            catch (UnauthorizedAccessException)
            {
                return new AppSettings();
            }
        }
    }

    public void Save(AppSettings settings)
    {
        lock (_sync)
        {
            SaveUnsafe(settings);
        }
    }

    private AppSettings CreateDefaultSettingsUnsafe()
    {
        var defaults = new AppSettings();
        SaveUnsafe(defaults);
        return defaults;
    }

    private AppSettings BackupAndResetSettingsUnsafe()
    {
        TryBackupCorruptedSettingsUnsafe();
        return CreateDefaultSettingsUnsafe();
    }

    private void SaveUnsafe(AppSettings settings)
    {
        Directory.CreateDirectory(BaseDirectory);

        var tempPath = SettingsPath + ".tmp";
        var json = JsonSerializer.Serialize(settings, SerializerOptions);
        File.WriteAllText(tempPath, json);
        File.Move(tempPath, SettingsPath, true);
    }

    private void TryBackupCorruptedSettingsUnsafe()
    {
        if (!File.Exists(SettingsPath))
        {
            return;
        }

        try
        {
            var backupPath = Path.Combine(BaseDirectory, $"settings.corrupt-{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.json");
            File.Copy(SettingsPath, backupPath);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
