using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

internal static class Program
{
    private const string MutexName = "Local\\SapB1ExcelPasteHelper.SingleInstance";

    [STAThread]
    private static void Main()
    {
        using var mutex = new Mutex(true, MutexName, out var isFirstInstance);
        if (!isFirstInstance)
        {
            MessageBox.Show(
                "SAP B1 Excel Helper is already running.",
                "SAP B1 Excel Helper",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        AppPaths.EnsureCreated();
        ApplicationConfiguration.Initialize();
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, args) =>
        {
            AppLogger.Error("UNHANDLED", args.Exception.Message, args.Exception);
            MessageBox.Show(args.Exception.Message, "Unexpected error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        };

        Application.Run(new MainForm());
        GC.KeepAlive(mutex);
    }
}

