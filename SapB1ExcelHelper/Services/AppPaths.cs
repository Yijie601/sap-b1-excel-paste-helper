namespace SapB1ExcelHelper.Services;

public static class AppPaths
{
    public static string DataDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SapB1ExcelHelper");

    public static string ConfigDirectory { get; } = Path.Combine(DataDirectory, "Config");
    public static string LogsDirectory { get; } = Path.Combine(DataDirectory, "Logs");
    public static string CalibrationFile { get; } = Path.Combine(ConfigDirectory, "calibration.json");
    public static string HotkeySettingsFile { get; } = Path.Combine(ConfigDirectory, "hotkey.json");
    public static string UpdateStateFile { get; } = Path.Combine(ConfigDirectory, "update_state.json");

    public static void EnsureCreated()
    {
        Directory.CreateDirectory(ConfigDirectory);
        Directory.CreateDirectory(LogsDirectory);
    }
}
