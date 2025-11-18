using TimeSheet_MAUI.Services;
using TimeSheet_MAUI.ViewModels;

namespace TimeSheet_MAUI;

public partial class MainPage : ContentPage
{
    private readonly MainViewModel _viewModel;
    private readonly IDialogService _dialogService;

    public MainPage(MainViewModel viewModel, IDialogService dialogService)
    {
        InitializeComponent();
        BindingContext = _viewModel = viewModel;
        _dialogService = dialogService;
    }

    protected override async void OnAppearing()
    {
        base.OnAppearing();
        _dialogService.Initialize(this);
        await _viewModel.InitializeAsync();
    }

}
