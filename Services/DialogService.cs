using System.IO;
using System.Linq;
using Microsoft.Maui.ApplicationModel;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Devices;
using Microsoft.Maui.Storage;
using TimeSheet_MAUI.ViewModels;
#if WINDOWS
using Microsoft.Maui.Platform;
using Windows.Storage.Pickers;
using WinRT.Interop;
#endif

namespace TimeSheet_MAUI.Services;

public sealed class DialogService : IDialogService
{
    private static readonly FilePickerFileType ExcelFileType = new(new Dictionary<DevicePlatform, IEnumerable<string>>
    {
        { DevicePlatform.WinUI, new[] { ".xlsx", ".xlsm", ".xls" } },
        { DevicePlatform.MacCatalyst, new[] { "com.microsoft.excel.xlsx", "org.openxmlformats.spreadsheetml.sheet" } },
        { DevicePlatform.iOS, new[] { "com.microsoft.excel.xlsx", "org.openxmlformats.spreadsheetml.sheet" } },
        { DevicePlatform.Android, new[] { "application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
        { DevicePlatform.Tizen, new[] { ".xlsx", ".xlsm", ".xls" } }
    });

    private Page? _page;

    public void Initialize(Page page) => _page = page;

    public async Task<string?> PickExcelFileAsync()
    {
        try
        {
            var file = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "Выберите Excel-файл",
                FileTypes = ExcelFileType
            });

            return file?.FullPath;
        }
        catch (Exception)
        {
            return null;
        }
    }

    public async Task ShowMessageAsync(string title, string message, StatusLevel level = StatusLevel.Info)
    {
        if (_page is null)
        {
            await MainThread.InvokeOnMainThreadAsync(() => Application.Current?.MainPage?.DisplayAlert(title, message, "Ок"));
            return;
        }

        await MainThread.InvokeOnMainThreadAsync(() => _page.DisplayAlert(title, message, "Ок"));
    }

    public async Task<bool> ShowConfirmationAsync(string title, string message, string acceptButton, string cancelButton)
    {
        return await MainThread.InvokeOnMainThreadAsync(async () =>
        {
            var targetPage = _page ?? Application.Current?.MainPage;
            if (targetPage is null)
            {
                return false;
            }

            return await targetPage.DisplayAlert(title, message, acceptButton, cancelButton);
        });
    }

    public async Task<string?> PickTemplateSavePathAsync(string suggestedFileName)
    {
#if WINDOWS
        try
        {
            var picker = new FileSavePicker
            {
                SuggestedFileName = Path.GetFileNameWithoutExtension(suggestedFileName)
            };
            picker.FileTypeChoices.Add("Excel", new[] { ".xlsx" });

            var window = Application.Current?.Windows.FirstOrDefault();
            if (window?.Handler?.PlatformView is MauiWinUIWindow nativeWindow)
            {
                InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(nativeWindow));
            }

            var file = await picker.PickSaveFileAsync();
            return file?.Path;
        }
        catch
        {
            return null;
        }
#else
        var directory = FileSystem.Current.AppDataDirectory;
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, suggestedFileName);
#endif
    }
}
