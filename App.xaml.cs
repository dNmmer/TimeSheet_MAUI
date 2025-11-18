using System;
using Microsoft.Maui;
#if WINDOWS
using Microsoft.Maui.Platform;
#endif

namespace TimeSheet_MAUI;

public partial class App : Application
{
    public App(AppShell shell)
    {
        InitializeComponent();
        MainPage = shell;
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
    private static void OnWindowHandlerChanged(object? sender, EventArgs e)
    {
        if (sender is Window window &&
            window.Handler?.PlatformView is MauiWinUIWindow mauiWindow)
        {
            mauiWindow.ExtendsContentIntoTitleBar = false;
        }
    }
#endif
}
