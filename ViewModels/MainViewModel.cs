using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Maui.ApplicationModel;
using Microsoft.Maui.Graphics;
using Microsoft.Maui.Storage;
using TimeSheet_MAUI.Services;
using Timesheet.Core.Configuration;
using Timesheet.Core.Models;
using Timesheet.Core.Services;

namespace TimeSheet_MAUI.ViewModels;

public partial class MainViewModel : ObservableRecipient
{
    private readonly ConfigService _configService;
    private readonly ExcelService _excelService;
    private readonly IDialogService _dialogs;
    private readonly ITimerService _timer;

    private AppConfig _currentConfig = AppConfig.Default;
    private bool _hasReference;
    private bool _pendingResumeChecked;
    private bool _isTimerPaused;
    private bool _shutdownHandlersRegistered;
    private int _shutdownPersisted;
    private CancellationTokenSource? _statusResetCts;

    public MainViewModel(
        ConfigService configService,
        ExcelService excelService,
        IDialogService dialogs,
        ITimerService timer)
    {
        _configService = configService;
        _excelService = excelService;
        _dialogs = dialogs;
        _timer = timer;
        _timer.Tick += OnTimerTick;

        BrowseFileCommand = new AsyncRelayCommand(BrowseFileAsync);
        ReloadReferenceCommand = new AsyncRelayCommand(ReloadReferenceAsync);
        CreateTemplateCommand = new AsyncRelayCommand(CreateTemplateAsync);
        ShowRequirementsCommand = new RelayCommand(ShowRequirements);
        ShowAboutCommand = new RelayCommand(ShowAbout);
        OpenCurrentFileCommand = new AsyncRelayCommand(OpenCurrentFileAsync);
        StartWorkdayCommand = new AsyncRelayCommand(StartWorkdayAsync, () => IsStartWorkdayEnabled);
        EndWorkdayCommand = new AsyncRelayCommand(EndWorkdayAsync, () => IsEndWorkdayEnabled);
        StartTimerCommand = new RelayCommand(StartTimer, CanStartTimer);
        PauseTimerCommand = new RelayCommand(PauseTimer, CanPauseTimer);
        StopTimerCommand = new AsyncRelayCommand(StopTimerAsync, CanStopTimer);

        ShowDefaultStatus();
        RegisterShutdownGuards();
    }

    public ObservableCollection<string> Projects { get; } = new();
    public ObservableCollection<string> WorkTypes { get; } = new();

    public IAsyncRelayCommand BrowseFileCommand { get; }
    public IAsyncRelayCommand ReloadReferenceCommand { get; }
    public IAsyncRelayCommand CreateTemplateCommand { get; }
    public IRelayCommand ShowRequirementsCommand { get; }
    public IRelayCommand ShowAboutCommand { get; }
    public IAsyncRelayCommand OpenCurrentFileCommand { get; }
    public IAsyncRelayCommand StartWorkdayCommand { get; }
    public IAsyncRelayCommand EndWorkdayCommand { get; }
    public IRelayCommand StartTimerCommand { get; }
    public IRelayCommand PauseTimerCommand { get; }
    public IAsyncRelayCommand StopTimerCommand { get; }

    [ObservableProperty]
    private string? excelPath;

    [ObservableProperty]
    private string? selectedProject;

    partial void OnSelectedProjectChanged(string? value)
    {
        NotifyTimerPrerequisitesChanged();
    }

    [ObservableProperty]
    private string? selectedWorkType;

    partial void OnSelectedWorkTypeChanged(string? value)
    {
        NotifyTimerPrerequisitesChanged();
    }

    [ObservableProperty]
    private string timerText = "00:00:00";

    [ObservableProperty]
    private bool isWorkdayStarted;

    [ObservableProperty]
    private string? statusMessage = "Файл не выбран.";

    [ObservableProperty]
    private bool isStatusOpen = true;

    [ObservableProperty]
    private bool isBusy;

    [ObservableProperty]
    private Color statusColor = Colors.Transparent;

    public bool AreInputsEnabled => _hasReference && !IsTimerActive && !IsBusy;

    public bool IsStartWorkdayEnabled => !IsTimerActive && !IsBusy && !IsWorkdayStarted && HasExcelPath;

    public bool IsEndWorkdayEnabled => !IsBusy && IsWorkdayStarted && !IsTimerActive;

    public bool IsStartButtonEnabled => CanStartTimer();

    public bool IsPauseButtonEnabled => _timer.IsRunning;

    public bool IsStopButtonEnabled => IsTimerActive && _timer.Elapsed > TimeSpan.Zero;

    private bool HasExcelPath => !string.IsNullOrWhiteSpace(ExcelPath);

    private bool IsTimerActive => _timer.IsRunning || _isTimerPaused;

    private bool _initialized;

    public async Task InitializeAsync()
    {
        if (_initialized)
        {
            return;
        }

        _initialized = true;

        _currentConfig = await _configService.LoadAsync();
        ExcelPath = _currentConfig.ExcelPath;

        if (HasExcelPath)
        {
            try
            {
                var loaded = await LoadReferenceAsync(ExcelPath!);
                if (loaded)
                {
                    await CheckPendingWorkdayAsync();
                }
            }
            catch (Exception ex)
            {
                await _dialogs.ShowMessageAsync("Ошибка", ex.Message, StatusLevel.Error);
            }
        }
    }

    private async Task BrowseFileAsync()
    {
        var selected = await _dialogs.PickExcelFileAsync();
        if (string.IsNullOrWhiteSpace(selected))
        {
            return;
        }

        var loaded = await LoadReferenceAsync(selected);
        if (loaded)
        {
            await SaveConfigAsync(selected);
        }
    }

    private async Task ReloadReferenceAsync()
    {
        if (!await EnsureExcelFileExistsAsync())
        {
            return;
        }

        await LoadReferenceAsync(ExcelPath!);
    }

    private async Task CreateTemplateAsync()
    {
        var suggestedName = $"timesheet_template_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        var savePath = await _dialogs.PickTemplateSavePathAsync(suggestedName);
        if (string.IsNullOrWhiteSpace(savePath))
        {
            return;
        }

        var created = await RunSafeAsync(async () =>
        {
            var directory = Path.GetDirectoryName(savePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            await Task.Run(() => _excelService.CreateTemplate(savePath));
        });

        if (!created)
        {
            return;
        }

        SetStatus("Откройте файл и заполните справочник.", StatusLevel.Info);
        var openedAutomatically = await TryOpenTemplateAsync(savePath);
        if (!openedAutomatically)
        {
            await _dialogs.ShowMessageAsync("Откройте файл", $"Откройте и заполните файл вручную:\n{savePath}", StatusLevel.Warning);
        }

        var opened = await WaitForFileOpenedAsync(savePath);
        if (!opened)
        {
            await _dialogs.ShowMessageAsync("Файл не заполнен", "Файл так и не был открыт. Повторите создание шаблона.", StatusLevel.Warning);
            return;
        }

        await WaitUntilFileReleasedAsync(savePath);

        var loaded = await LoadReferenceAsync(savePath);
        if (loaded)
        {
            await SaveConfigAsync(savePath);
        }
    }

        
        
    
    
    
        private void ShowRequirements()
    {
        var message =
            "Файл должен содержать листы:\n" +
            "• 'Справочник' со столбцами Проект и Вид работ\n" +
            "• 'Учет времени' со столбцами Дата, Проект, Вид работ, Длительность, В часах, Комментарий\n" +
            "• 'Учет рабочего времени' со столбцами Дата, Начало, Окончание, Длительность, В часах";
        _ = _dialogs.ShowMessageAsync("Требования к Excel", message);
    }

    private void ShowAbout()
    {
        var message = $"Timesheet .NET MAUI\nВерсия {typeof(MainViewModel).Assembly.GetName().Version}";
        _ = _dialogs.ShowMessageAsync("О приложении", message);
    }

    private async Task OpenCurrentFileAsync()
    {
        if (!HasExcelPath)
        {
            await _dialogs.ShowMessageAsync("Файл не выбран", "Укажите Excel-файл.", StatusLevel.Warning);
            return;
        }

        if (!File.Exists(ExcelPath))
        {
            await _dialogs.ShowMessageAsync("Файл не найден", "Указанный Excel-файл отсутствует.", StatusLevel.Error);
            return;
        }

        try
        {
            var file = new ReadOnlyFile(ExcelPath!);
            await Launcher.Default.OpenAsync(new OpenFileRequest(Path.GetFileName(ExcelPath), file));
        }
        catch (Exception ex)
        {
            await _dialogs.ShowMessageAsync("Не удалось открыть файл", ex.Message, StatusLevel.Error);
        }
    }

    private async Task StartWorkdayAsync()
    {
        if (!await EnsureExcelFileExistsAsync())
        {
            return;
        }

        await RunSafeAsync(async () =>
        {
            var info = await Task.Run(() => _excelService.StartWorkday(ExcelPath!));
            IsWorkdayStarted = true;
            SetStatus($"Рабочий день начат в {info.StartTime}.", StatusLevel.Success);
        });
    }

    private async Task EndWorkdayAsync()
    {
        if (!await EnsureExcelFileExistsAsync())
        {
            return;
        }

        if (!IsWorkdayStarted)
        {
            await _dialogs.ShowMessageAsync("Рабочий день ещё не начат", "Сначала начните рабочий день.", StatusLevel.Warning);
            return;
        }

        await RunSafeAsync(async () =>
        {
            var info = await Task.Run(() => _excelService.EndWorkday(ExcelPath!));
            IsWorkdayStarted = false;
            SetStatus($"Рабочий день завершён. Длительность: {info.Duration}.", StatusLevel.Success);
        });
    }

    private void StartTimer()
    {
        _timer.Start(_timer.Elapsed);
        _isTimerPaused = false;
        NotifyInteractionStateChanged();
    }

    private void PauseTimer()
    {
        if (!_timer.IsRunning)
        {
            return;
        }

        _timer.Pause();
        _isTimerPaused = true;
        NotifyInteractionStateChanged();
    }

    private async Task StopTimerAsync()
    {
        var elapsed = _timer.Elapsed;
        _timer.Stop();
        _isTimerPaused = false;
        NotifyInteractionStateChanged();

        if (elapsed <= TimeSpan.Zero)
        {
            return;
        }

        if (!await EnsureExcelFileExistsAsync())
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(SelectedProject) || string.IsNullOrWhiteSpace(SelectedWorkType))
        {
            await _dialogs.ShowMessageAsync("Недостаточно данных", "Выберите проект и тип работ.", StatusLevel.Warning);
            return;
        }

        await RunSafeAsync(async () =>
        {
            string? comment = null;
            var addComment = await _dialogs.ShowConfirmationAsync("Оставить комментарий?", string.Empty, "Да", "Нет");
            if (addComment)
            {
                comment = await _dialogs.ShowPromptAsync("Комментарий", "Введите текст комментария", "Ок", "Отмена");
                if (string.IsNullOrWhiteSpace(comment))
                {
                    comment = null;
                }
            }

            var entry = new TimesheetEntry(SelectedProject!, SelectedWorkType!, elapsed, DateTime.Now, comment);
            await Task.Run(() => _excelService.AppendTimeEntry(ExcelPath!, entry));
            SetStatus("Запись добавлена в 'Учет времени'.", StatusLevel.Success);
            TimerText = FormatElapsed(TimeSpan.Zero);
        });
    }

    private bool CanStartTimer()
    {
        if (IsBusy || !IsWorkdayStarted || !HasExcelPath)
        {
            return false;
        }

        if (_timer.IsRunning)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(SelectedProject) || string.IsNullOrWhiteSpace(SelectedWorkType))
        {
            return false;
        }

        return _isTimerPaused || AreInputsEnabled;
    }

    private bool CanPauseTimer() => _timer.IsRunning;

    private bool CanStopTimer() => _timer.Elapsed > TimeSpan.Zero && IsTimerActive;

    private async Task<bool> LoadReferenceAsync(string path)
    {
        while (true)
        {
            if (!File.Exists(path))
            {
                await HandleMissingExcelFileAsync(path);
                return false;
            }

            string? structureErrorMessage = null;

            var success = await RunSafeAsync(async () =>
            {
                try
                {
                    var data = await Task.Run(() => _excelService.LoadReferenceData(path));
                    ReplaceItems(Projects, data.Projects);
                    ReplaceItems(WorkTypes, data.WorkTypes);
                    _hasReference = Projects.Count > 0 && WorkTypes.Count > 0;
                    if (!_hasReference)
                    {
                        throw new ExcelStructureException("Лист 'Справочник' должен содержать хотя бы один проект и тип работ.");
                    }

                    SelectedProject = Projects.FirstOrDefault();
                    SelectedWorkType = WorkTypes.FirstOrDefault();
                    ExcelPath = path;
                    SetStatus($"Выбран файл: {path}", StatusLevel.Success);
                }
                catch (ExcelStructureException ex)
                {
                    structureErrorMessage = ex.Message;
                    throw;
                }
            }, showStructureError: false);

            if (success)
            {
                return true;
            }

            if (structureErrorMessage is null)
            {
                ShowDefaultStatus();
                return false;
            }

            var shouldOpen = await _dialogs.ShowConfirmationAsync(
                "Справочник не заполнен",
                $"{structureErrorMessage}\nОткрыть файл для заполнения?",
                "Заполнить",
                "Ок");

            if (!shouldOpen)
            {
                ShowDefaultStatus();
                return false;
            }

            var opened = await TryOpenTemplateAsync(path);
            if (!opened)
            {
                await _dialogs.ShowMessageAsync("Не удалось открыть файл", $"Откройте и заполните файл вручную:\n{path}", StatusLevel.Error);
                continue;
            }

            var hasBeenOpened = await WaitForFileOpenedAsync(path);
            if (!hasBeenOpened)
            {
                await _dialogs.ShowMessageAsync("Файл не открыт", "Файл не был открыт. Повторите попытку.", StatusLevel.Warning);
                continue;
            }

            await WaitUntilFileReleasedAsync(path);
        }
    }

    private async Task SaveConfigAsync(string path)
    {
        _currentConfig = new AppConfig(path);
        await _configService.SaveAsync(_currentConfig);
    }

    private async Task<bool> RunSafeAsync(Func<Task> action, bool showStructureError = true)
    {
        if (IsBusy)
        {
            return false;
        }

        try
        {
            IsBusy = true;
            NotifyInteractionStateChanged();
            await action();
            return true;
        }
        catch (ExcelStructureException ex)
        {
            if (showStructureError)
            {
                await _dialogs.ShowMessageAsync("Структура Excel", ex.Message, StatusLevel.Error);
            }
            return false;
        }
        catch (Exception ex)
        {
            await _dialogs.ShowMessageAsync("Ошибка", ex.Message, StatusLevel.Error);
            return false;
        }
        finally
        {
            IsBusy = false;
            NotifyInteractionStateChanged();
        }
    }

    private void OnTimerTick(object? sender, TimeSpan elapsed)
    {
        TimerText = FormatElapsed(elapsed);
        NotifyInteractionStateChanged();
    }

    private async Task<bool> TryOpenTemplateAsync(string path)
    {
        try
        {
#if WINDOWS
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            await Task.CompletedTask;
            return true;
#else
            var file = new ReadOnlyFile(path);
            await Launcher.Default.OpenAsync(new OpenFileRequest(Path.GetFileName(path), file));
            return true;
#endif
        }
        catch
        {
            return false;
        }
    }

    private async Task<bool> WaitForFileOpenedAsync(string path)
    {
        var timeout = TimeSpan.FromMinutes(5);
        var watch = Stopwatch.StartNew();

        while (watch.Elapsed < timeout)
        {
            if (!File.Exists(path))
            {
                return false;
            }

            if (IsFileLocked(path))
            {
                return true;
            }

            await Task.Delay(TimeSpan.FromSeconds(1));
        }

        return false;
    }

    private async Task WaitUntilFileReleasedAsync(string path)
    {
        while (File.Exists(path) && IsFileLocked(path))
        {
            await Task.Delay(TimeSpan.FromSeconds(1));
        }
    }

    private static bool IsFileLocked(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
    }

    private async Task CheckPendingWorkdayAsync()
    {
        if (_pendingResumeChecked || !HasExcelPath)
        {
            return;
        }

        _pendingResumeChecked = true;

        WorkdayPendingInfo? info;
        try
        {
            info = await Task.Run(() => _excelService.GetPendingWorkday(ExcelPath!, DateTime.Today));
        }
        catch (ExcelStructureException ex)
        {
            await _dialogs.ShowMessageAsync("Структура Excel", ex.Message, StatusLevel.Error);
            return;
        }
        catch (Exception ex)
        {
            await _dialogs.ShowMessageAsync("Ошибка", ex.Message, StatusLevel.Error);
            return;
        }

        if (info is null)
        {
            return;
        }

        var resume = await _dialogs.ShowConfirmationAsync(
            "Продолжить рабочий день?",
            $"Рабочий день за {info.Date:dd.MM.yyyy} начат в {info.StartTime:hh\\:mm}. Продолжить?",
            "Да",
            "Нет");

        if (!resume)
        {
            return;
        }

        IsWorkdayStarted = true;
        SetStatus($"Рабочий день продолжается с {info.StartTime:hh\\:mm}.", StatusLevel.Info);
        NotifyInteractionStateChanged();
    }

    public async Task<bool> ConfirmCloseAsync()
    {
        if (IsTimerActive)
        {
            await _dialogs.ShowMessageAsync("Таймер активен", "Остановите таймер перед выходом.", StatusLevel.Warning);
            return false;
        }

        if (IsWorkdayStarted)
        {
            await _dialogs.ShowMessageAsync("Рабочий день не завершён", "Завершите рабочий день перед выходом.", StatusLevel.Warning);
            return false;
        }

        return true;
    }

    public async Task ForcePersistStateAsync()
    {
        if (Interlocked.Exchange(ref _shutdownPersisted, 1) == 1)
        {
            return;
        }

        if (!HasExcelPath)
        {
            return;
        }

        var path = ExcelPath!;

        try
        {
            if (IsTimerActive &&
                _timer.Elapsed > TimeSpan.Zero &&
                !string.IsNullOrWhiteSpace(SelectedProject) &&
                !string.IsNullOrWhiteSpace(SelectedWorkType))
            {
                var entry = new TimesheetEntry(SelectedProject!, SelectedWorkType!, _timer.Elapsed, DateTime.Now);
                await Task.Run(() => _excelService.AppendTimeEntry(path, entry));
            }

            if (IsWorkdayStarted)
            {
                await Task.Run(() => _excelService.EndWorkday(path));
                IsWorkdayStarted = false;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Shutdown] {ex}");
        }
        finally
        {
            _timer.Stop();
            _isTimerPaused = false;
        }
    }

    private void RegisterShutdownGuards()
    {
        if (_shutdownHandlersRegistered)
        {
            return;
        }

        AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        _shutdownHandlersRegistered = true;
    }

    private void OnProcessExit(object? sender, EventArgs e) => ForcePersistStateSync();

    private void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e) => ForcePersistStateSync();

    private void ForcePersistStateSync()
    {
        try
        {
            ForcePersistStateAsync().GetAwaiter().GetResult();
        }
        catch
        {
            // ignored
        }
    }

    private async Task<bool> EnsureExcelFileExistsAsync()
    {
        if (!HasExcelPath)
        {
            await _dialogs.ShowMessageAsync("Файл не выбран", "Укажите Excel-файл.", StatusLevel.Warning);
            return false;
        }

        if (!File.Exists(ExcelPath!))
        {
            await HandleMissingExcelFileAsync(ExcelPath!);
            return false;
        }

        return true;
    }

    private async Task HandleMissingExcelFileAsync(string path)
    {
        await _dialogs.ShowMessageAsync("Файл не найден", $"Файл недоступен:\n{path}", StatusLevel.Error);

        if (string.Equals(path, ExcelPath, StringComparison.OrdinalIgnoreCase))
        {
            ExcelPath = null;
            _hasReference = false;
            Projects.Clear();
            WorkTypes.Clear();
            SelectedProject = null;
            SelectedWorkType = null;
            _currentConfig = AppConfig.Default;
            await _configService.SaveAsync(_currentConfig);
            ShowDefaultStatus();
        }
    }

    private static void ReplaceItems(ObservableCollection<string> collection, IReadOnlyList<string> items)
    {
        collection.Clear();
        foreach (var item in items)
        {
            collection.Add(item);
        }
    }

    private static string FormatElapsed(TimeSpan elapsed) =>
        $"{(int)elapsed.TotalHours:00}:{elapsed.Minutes:00}:{elapsed.Seconds:00}";

    private void SetStatus(string message, StatusLevel level)
    {
        StatusMessage = message;
        IsStatusOpen = true;
        StatusColor = level switch
        {
            StatusLevel.Success => Color.FromArgb("#1C8F36"),
            StatusLevel.Warning => Color.FromArgb("#E3A018"),
            StatusLevel.Error => Color.FromArgb("#C42B1C"),
            _ => Color.FromArgb("#2563EB")
        };
        ScheduleStatusReset();
    }

    private void NotifyInteractionStateChanged()
    {
        OnPropertyChanged(nameof(AreInputsEnabled));
        OnPropertyChanged(nameof(IsStartWorkdayEnabled));
        OnPropertyChanged(nameof(IsEndWorkdayEnabled));
        OnPropertyChanged(nameof(IsStartButtonEnabled));
        OnPropertyChanged(nameof(IsPauseButtonEnabled));
        OnPropertyChanged(nameof(IsStopButtonEnabled));
        StartWorkdayCommand.NotifyCanExecuteChanged();
        EndWorkdayCommand.NotifyCanExecuteChanged();
        StartTimerCommand.NotifyCanExecuteChanged();
        PauseTimerCommand.NotifyCanExecuteChanged();
        StopTimerCommand.NotifyCanExecuteChanged();
    }

    private void NotifyTimerPrerequisitesChanged()
    {
        StartTimerCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(IsStartButtonEnabled));
    }

    private void ScheduleStatusReset()
    {
        _statusResetCts?.Cancel();
        var cts = new CancellationTokenSource();
        _statusResetCts = cts;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(10), cts.Token);
                await MainThread.InvokeOnMainThreadAsync(ShowDefaultStatus);
            }
            catch (OperationCanceledException)
            {
                // ignored
            }
        });
    }

    private void ShowDefaultStatus()
    {
        _statusResetCts?.Cancel();
        StatusColor = Colors.Transparent;
        StatusMessage = GetDefaultStatusMessage();
        IsStatusOpen = true;
    }

    private string GetDefaultStatusMessage() =>
        HasExcelPath ? $"Выбран файл: {ExcelPath}" : "Файл не выбран.";
}


