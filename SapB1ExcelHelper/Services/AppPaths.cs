namespace SapB1ExcelHelper.Services;

public static class AppPaths
{
    public static string DataDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SapB1ExcelHelper");

    public static string ConfigDirectory { get; } = Path.Combine(DataDirectory, "Config");
    public static string LogsDirectory { get; } = Path.Combine(DataDirectory, "Logs");
    public static string SupplierMappingFile { get; } = Path.Combine(ConfigDirectory, "supplier_mapping.csv");
    public static string CalibrationFile { get; } = Path.Combine(ConfigDirectory, "calibration.json");
    public static string HotkeySettingsFile { get; } = Path.Combine(ConfigDirectory, "hotkey.json");
    public static string UpdateStateFile { get; } = Path.Combine(ConfigDirectory, "update_state.json");

    public static void EnsureCreated()
    {
        Directory.CreateDirectory(ConfigDirectory);
        Directory.CreateDirectory(LogsDirectory);

        if (!File.Exists(SupplierMappingFile))
        {
            var packagedFile = Path.Combine(AppContext.BaseDirectory, "Config", "supplier_mapping.csv");
            if (File.Exists(packagedFile))
            {
                File.Copy(packagedFile, SupplierMappingFile);
            }
            else
            {
                File.WriteAllText(SupplierMappingFile, "Supplier Name,SAP Code\r\n");
            }
        }
    }
}
