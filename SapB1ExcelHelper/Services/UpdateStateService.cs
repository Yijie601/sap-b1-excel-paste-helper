using System.Text.Json;

namespace SapB1ExcelHelper.Services;

public sealed class UpdateStateService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public bool IsAutomaticCheckDue(TimeSpan interval)
    {
        try
        {
            if (!File.Exists(AppPaths.UpdateStateFile))
            {
                return true;
            }

            var state = JsonSerializer.Deserialize<UpdateState>(
                File.ReadAllText(AppPaths.UpdateStateFile),
                JsonOptions);
            return state?.LastCheckedUtc is null ||
                   DateTimeOffset.UtcNow - state.LastCheckedUtc.Value >= interval;
        }
        catch
        {
            return true;
        }
    }

    public void RecordCheck()
    {
        try
        {
            var temporaryFile = AppPaths.UpdateStateFile + ".tmp";
            File.WriteAllText(temporaryFile, JsonSerializer.Serialize(
                new UpdateState { LastCheckedUtc = DateTimeOffset.UtcNow },
                JsonOptions));
            File.Move(temporaryFile, AppPaths.UpdateStateFile, true);
        }
        catch (Exception exception)
        {
            AppLogger.Error("UPDATE_STATE_ERROR", exception.Message, exception);
        }
    }

    private sealed class UpdateState
    {
        public DateTimeOffset? LastCheckedUtc { get; init; }
    }
}
