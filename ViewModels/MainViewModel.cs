using System.Collections.ObjectModel;
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

    [ObservableProperty]
    private string? selectedWorkType;

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

    public bool AreInputsEnabled => _hasReference && !_timer.IsRunning && !IsBusy;

    public bool IsStartWorkdayEnabled => !_timer.IsRunning && !IsBusy && !IsWorkdayStarted && HasExcelPath;

    public bool IsEndWorkdayEnabled => !IsBusy && IsWorkdayStarted;

    private bool HasExcelPath => !string.IsNullOrWhiteSpace(ExcelPath);

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
                await LoadReferenceAsync(ExcelPath!);
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

        await LoadReferenceAsync(selected);
        await SaveConfigAsync(selected);
    }

    private async Task ReloadReferenceAsync()
    {
        if (!HasExcelPath)
        {
            await _dialogs.ShowMessageAsync("Файл не выбран", "Укажите Excel-файл.", StatusLevel.Warning);
            return;
        }

        await LoadReferenceAsync(ExcelPath!);
    }

    private async Task CreateTemplateAsync()
    {
        var directory = FileSystem.Current.AppDataDirectory;
        Directory.CreateDirectory(directory);
        var fileName = $"timesheet_template_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        var savePath = Path.Combine(directory, fileName);

        await RunSafeAsync(async () =>
        {
            await Task.Run(() => _excelService.CreateTemplate(savePath));
            await _dialogs.ShowMessageAsync("Готово", $"Шаблон создан:\n{savePath}");
            await LoadReferenceAsync(savePath);
            await SaveConfigAsync(savePath);
        });
    }

    private void ShowRequirements()
    {
        var message =
            "Файл должен содержать листы:\n" +
            "• 'Справочник' со столбцами Проект и Тип работ\n" +
            "• 'Учет времени' со столбцами Дата, Проект, Тип работ, Длительность\n" +
            "• 'Учет рабочего времени' со столбцами Дата, Начало, Окончание, Длительность";
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
        if (!HasExcelPath)
        {
            await _dialogs.ShowMessageAsync("Файл не выбран", "Укажите Excel-файл.", StatusLevel.Warning);
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
        if (!HasExcelPath)
        {
            await _dialogs.ShowMessageAsync("Файл не выбран", "Укажите Excel-файл.", StatusLevel.Warning);
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
        NotifyInteractionStateChanged();
    }

    private void PauseTimer()
    {
        _timer.Pause();
        NotifyInteractionStateChanged();
    }

    private async Task StopTimerAsync()
    {
        var elapsed = _timer.Elapsed;
        _timer.Stop();
        NotifyInteractionStateChanged();

        if (elapsed <= TimeSpan.Zero)
        {
            return;
        }

        if (!HasExcelPath)
        {
            await _dialogs.ShowMessageAsync("Файл не выбран", "Укажите Excel-файл.", StatusLevel.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(SelectedProject) || string.IsNullOrWhiteSpace(SelectedWorkType))
        {
            await _dialogs.ShowMessageAsync("Недостаточно данных", "Выберите проект и тип работ.", StatusLevel.Warning);
            return;
        }

        await RunSafeAsync(async () =>
        {
            var entry = new TimesheetEntry(SelectedProject!, SelectedWorkType!, elapsed, DateTime.Now);
            await Task.Run(() => _excelService.AppendTimeEntry(ExcelPath!, entry));
            SetStatus("Запись добавлена в 'Учет времени'.", StatusLevel.Success);
            TimerText = FormatElapsed(TimeSpan.Zero);
        });
    }

    private bool CanStartTimer() =>
        IsWorkdayStarted && AreInputsEnabled && !string.IsNullOrWhiteSpace(SelectedProject) && !string.IsNullOrWhiteSpace(SelectedWorkType);

    private bool CanPauseTimer() => _timer.IsRunning;

    private bool CanStopTimer() => _timer.Elapsed > TimeSpan.Zero;

    private async Task LoadReferenceAsync(string path)
    {
        await RunSafeAsync(async () =>
        {
            var data = await Task.Run(() => _excelService.LoadReferenceData(path));
            ReplaceItems(Projects, data.Projects);
            ReplaceItems(WorkTypes, data.WorkTypes);
            SelectedProject = Projects.FirstOrDefault();
            SelectedWorkType = WorkTypes.FirstOrDefault();
            _hasReference = Projects.Count > 0 && WorkTypes.Count > 0;
            ExcelPath = path;
            SetStatus($"Выбран файл: {path}", StatusLevel.Success);
        });
    }

    private async Task SaveConfigAsync(string path)
    {
        _currentConfig = new AppConfig(path);
        await _configService.SaveAsync(_currentConfig);
    }

    private async Task RunSafeAsync(Func<Task> action)
    {
        if (IsBusy)
        {
            return;
        }

        try
        {
            IsBusy = true;
            NotifyInteractionStateChanged();
            await action();
        }
        catch (ExcelStructureException ex)
        {
            await _dialogs.ShowMessageAsync("Структура Excel", ex.Message, StatusLevel.Error);
        }
        catch (Exception ex)
        {
            await _dialogs.ShowMessageAsync("Ошибка", ex.Message, StatusLevel.Error);
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
    }

    private void NotifyInteractionStateChanged()
    {
        OnPropertyChanged(nameof(AreInputsEnabled));
        OnPropertyChanged(nameof(IsStartWorkdayEnabled));
        OnPropertyChanged(nameof(IsEndWorkdayEnabled));
        StartWorkdayCommand.NotifyCanExecuteChanged();
        EndWorkdayCommand.NotifyCanExecuteChanged();
        StartTimerCommand.NotifyCanExecuteChanged();
        PauseTimerCommand.NotifyCanExecuteChanged();
        StopTimerCommand.NotifyCanExecuteChanged();
    }
}
