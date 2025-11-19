using System;
using System.Threading.Tasks;
using Microsoft.Maui;
#if WINDOWS
using Microsoft.Maui.ApplicationModel;
using Microsoft.Maui.Platform;
using Microsoft.UI.Windowing;
#endif
using TimeSheet_MAUI.ViewModels;

namespace TimeSheet_MAUI;

public partial class App : Application
{
    private readonly MainViewModel _viewModel;
#if WINDOWS
    private Window? _currentWindow;
    private AppWindow? _appWindow;
    private bool _programmaticClose;
#endif

    public App(AppShell shell, MainViewModel viewModel)
    {
        InitializeComponent();
        MainPage = shell;
        _viewModel = viewModel;
    }

    protected override Window CreateWindow(IActivationState? activationState)
    {
        var window = base.CreateWindow(activationState);
        window.Title = "TimeSheet";
        window.Width = 720;
        window.Height = 520;
        window.MinimumWidth = 600;
        window.MinimumHeight = 420;
#if WINDOWS
        window.HandlerChanged += OnWindowHandlerChanged;
#endif
        return window;
    }

#if WINDOWS
    private void OnWindowHandlerChanged(object? sender, EventArgs e)
    {
        if (sender is not Window window ||
            window.Handler?.PlatformView is not MauiWinUIWindow mauiWindow)
        {
            return;
        }

        mauiWindow.ExtendsContentIntoTitleBar = false;
        _currentWindow = window;
        AttachClosingHandler(mauiWindow);
    }

    private void AttachClosingHandler(MauiWinUIWindow mauiWindow)
    {
        if (_appWindow is not null)
        {
            _appWindow.Closing -= OnAppWindowClosing;
        }

        _appWindow = mauiWindow.AppWindow;
        _appWindow.Closing += OnAppWindowClosing;
    }

    private void OnAppWindowClosing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_viewModel is null)
        {
            return;
        }

        if (_programmaticClose)
        {
            _programmaticClose = false;
            return;
        }

        args.Cancel = true;
        _ = HandleCloseRequestAsync();
    }

    private async Task HandleCloseRequestAsync()
    {
        if (_viewModel is null)
        {
            return;
        }

        var canClose = await _viewModel.ConfirmCloseAsync();
        if (!canClose)
        {
            return;
        }

        await _viewModel.ForcePersistStateAsync();
        _programmaticClose = true;
        await MainThread.InvokeOnMainThreadAsync(() =>
        {
            if (_currentWindow is not null)
            {
                Application.Current?.CloseWindow(_currentWindow);
            }
        });
    }
#endif
}
