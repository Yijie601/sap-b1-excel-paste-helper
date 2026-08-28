using System.Text.Json;
using SapB1ExcelHelper.Models;

namespace SapB1ExcelHelper.Services;

public sealed class CalibrationService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public SapCalibration Load()
    {
        try
        {
            if (File.Exists(AppPaths.CalibrationFile))
            {
                return JsonSerializer.Deserialize<SapCalibration>(
                           File.ReadAllText(AppPaths.CalibrationFile),
                           JsonOptions) ?? new SapCalibration();
            }
        }
        catch (Exception exception)
        {
            AppLogger.Error("CALIBRATION_LOAD_ERROR", exception.Message, exception);
        }

        var defaults = new SapCalibration();
        Save(defaults);
        return defaults;
    }

    public void Save(SapCalibration calibration)
    {
        Directory.CreateDirectory(AppPaths.ConfigDirectory);
        var temporaryFile = AppPaths.CalibrationFile + ".tmp";
        File.WriteAllText(temporaryFile, JsonSerializer.Serialize(calibration, JsonOptions));
        File.Move(temporaryFile, AppPaths.CalibrationFile, true);
    }
}
