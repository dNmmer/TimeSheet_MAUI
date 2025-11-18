using System.IO;
using Microsoft.Extensions.Logging;
using TimeSheet_MAUI.Services;
using TimeSheet_MAUI.ViewModels;
using Timesheet.Core.Configuration;
using Timesheet.Core.Services;

namespace TimeSheet_MAUI;

public static class MauiProgram
{
    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder
            .UseMauiApp<App>()
            .ConfigureFonts(fonts =>
            {
                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

        builder.Services.AddSingleton<AppShell>();
        builder.Services.AddSingleton<MainPage>();
        builder.Services.AddSingleton<MainViewModel>();

        builder.Services.AddSingleton<ConfigService>();
        builder.Services.AddSingleton<ExcelService>();
        builder.Services.AddSingleton<IDialogService, DialogService>();
        builder.Services.AddSingleton<ITimerService, DispatcherTimerService>();

#if DEBUG
        builder.Logging.AddDebug();
#endif

        RegisterGlobalExceptionHandlers();

        return builder.Build();
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        var logPath = Path.Combine(AppContext.BaseDirectory, "app-crash.log");
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            try
            {
                File.AppendAllText(logPath,
                    $"[{DateTime.Now:O}] Unhandled: {args.ExceptionObject}\n");
            }
            catch
            {
                // ignore logging failures
            }
        };

        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            try
            {
                File.AppendAllText(logPath,
                    $"[{DateTime.Now:O}] Unobserved: {args.Exception}\n");
            }
            catch
            {
                // ignore logging failures
            }
        };
    }
}
