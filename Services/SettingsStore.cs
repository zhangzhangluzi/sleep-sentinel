using System.Text.Json;
using System.Text;
using SleepSentinel.Models;

namespace SleepSentinel.Services;

public sealed class SettingsStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true
    };
    private readonly object _sync = new();
    private AppSettings? _cachedSettings;

    public string BaseDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SleepSentinel");

    public string SettingsPath => Path.Combine(BaseDirectory, "settings.json");
    public string LastKnownGoodSettingsPath => Path.Combine(BaseDirectory, "settings.last-good.json");

    public AppSettings Load()
    {
        lock (_sync)
        {
            var settings = LoadUnsafe();
            _cachedSettings = settings.Clone();
            return settings.Clone();
        }
    }

    public void Save(AppSettings settings)
    {
        lock (_sync)
        {
            SaveUnsafe(settings);
        }
    }

    public AppSettings Update(Action<AppSettings> mutation)
    {
        lock (_sync)
        {
            var settings = _cachedSettings?.Clone() ?? LoadUnsafe();
            mutation(settings);
            SaveUnsafe(settings);
            return settings.Clone();
        }
    }

    private AppSettings LoadUnsafe()
    {
        Directory.CreateDirectory(BaseDirectory);

        if (!File.Exists(SettingsPath))
        {
            if (TryLoadLastKnownGoodUnsafe(out var recovered))
            {
                SaveUnsafe(recovered);
                return recovered;
            }

            return CreateDefaultSettingsUnsafe();
        }

        try
        {
            return ReadSettingsFileUnsafe(SettingsPath) ?? BackupAndResetSettingsUnsafe();
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
            if (_cachedSettings is not null)
            {
                return _cachedSettings.Clone();
            }

            if (TryLoadLastKnownGoodUnsafe(out var recovered))
            {
                return recovered;
            }

            return new AppSettings();
        }
        catch (UnauthorizedAccessException)
        {
            if (_cachedSettings is not null)
            {
                return _cachedSettings.Clone();
            }

            if (TryLoadLastKnownGoodUnsafe(out var recovered))
            {
                return recovered;
            }

            return new AppSettings();
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
        try
        {
            File.WriteAllText(tempPath, json, Encoding.UTF8);
            File.Move(tempPath, SettingsPath, true);
        }
        finally
        {
            TryDeleteTempFile(tempPath);
        }

        TryWriteLastKnownGoodUnsafe(json);
        _cachedSettings = settings.Clone();
    }

    private AppSettings? ReadSettingsFileUnsafe(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<AppSettings>(json, SerializerOptions);
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

    private bool TryLoadLastKnownGoodUnsafe(out AppSettings settings)
    {
        settings = new AppSettings();
        if (!File.Exists(LastKnownGoodSettingsPath))
        {
            return false;
        }

        try
        {
            settings = ReadSettingsFileUnsafe(LastKnownGoodSettingsPath) ?? new AppSettings();
            return true;
        }
        catch (JsonException)
        {
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
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

    private void TryWriteLastKnownGoodUnsafe(string json)
    {
        var tempPath = LastKnownGoodSettingsPath + ".tmp";
        try
        {
            File.WriteAllText(tempPath, json, Encoding.UTF8);
            File.Move(tempPath, LastKnownGoodSettingsPath, true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
        finally
        {
            TryDeleteTempFile(tempPath);
        }
    }

    private static void TryDeleteTempFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
