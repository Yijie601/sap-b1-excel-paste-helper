using System.Text.Json;

namespace SapB1ExcelHelper.Services;

public sealed class HotkeySettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    private readonly string _filePath;

    public HotkeySettingsService(string? filePath = null)
    {
        _filePath = filePath ?? AppPaths.HotkeySettingsFile;
    }

    public HotkeyDefinition Load()
    {
        try
        {
            if (!File.Exists(_filePath))
            {
                return HotkeyDefinition.Default;
            }

            var hotkey = JsonSerializer.Deserialize<HotkeyDefinition>(
                File.ReadAllText(_filePath),
                JsonOptions);
            return hotkey is not null && hotkey.IsSupported(out _)
                ? hotkey
                : HotkeyDefinition.Default;
        }
        catch (Exception exception)
        {
            AppLogger.Error("HOTKEY_LOAD_ERROR", exception.Message, exception);
            return HotkeyDefinition.Default;
        }
    }

    public void Save(HotkeyDefinition hotkey)
    {
        if (!hotkey.IsSupported(out var error))
        {
            throw new ArgumentException(error, nameof(hotkey));
        }

        var directory = Path.GetDirectoryName(_filePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var temporaryFile = _filePath + ".tmp";
        File.WriteAllText(temporaryFile, JsonSerializer.Serialize(hotkey, JsonOptions));
        File.Move(temporaryFile, _filePath, true);
    }
}
