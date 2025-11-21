using TimeSheet_MAUI.ViewModels;

namespace TimeSheet_MAUI.Services;

public interface IDialogService
{
    void Initialize(Page page);

    Task<string?> PickExcelFileAsync();

    Task<string?> PickTemplateSavePathAsync(string suggestedFileName);

    Task ShowMessageAsync(string title, string message, StatusLevel level = StatusLevel.Info);

    Task<bool> ShowConfirmationAsync(string title, string message, string acceptButton, string cancelButton);

    Task<string?> ShowPromptAsync(string title, string message, string acceptButton, string cancelButton);
}

