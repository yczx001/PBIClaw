using System.Text.Json;

namespace PbiMetadataTool;

internal sealed class AbiSettingsStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    private readonly string _settingsPath;

    public AbiSettingsStore(string? settingsPath = null)
    {
        _settingsPath = settingsPath ?? BuildDefaultPath();
    }

    public AbiAssistantSettings Load()
    {
        try
        {
            if (!File.Exists(_settingsPath))
            {
                return new AbiAssistantSettings();
            }

            var json = File.ReadAllText(_settingsPath);
            return JsonSerializer.Deserialize<AbiAssistantSettings>(json, SerializerOptions) ?? new AbiAssistantSettings();
        }
        catch
        {
            return new AbiAssistantSettings();
        }
    }

    public void Save(AbiAssistantSettings settings)
    {
        var directory = Path.GetDirectoryName(_settingsPath);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(settings, SerializerOptions);
        File.WriteAllText(_settingsPath, json);
    }

    private static string BuildDefaultPath()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        return Path.Combine(appData, "PBIClaw", "settings.json");
    }
}
