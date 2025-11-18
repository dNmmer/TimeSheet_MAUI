using Microsoft.Maui.ApplicationModel;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Devices;
using Microsoft.Maui.Storage;
using TimeSheet_MAUI.ViewModels;

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
}
