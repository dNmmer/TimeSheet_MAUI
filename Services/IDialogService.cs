using TimeSheet_MAUI.ViewModels;

namespace TimeSheet_MAUI.Services;

public interface IDialogService
{
    void Initialize(Page page);

    Task<string?> PickExcelFileAsync();

    Task ShowMessageAsync(string title, string message, StatusLevel level = StatusLevel.Info);
}
