using Microsoft.Maui.ApplicationModel;

namespace TimeSheet_MAUI.Services;

public sealed class DispatcherTimerService : ITimerService
{
    private readonly System.Timers.Timer _timer;
    private TimeSpan _offset = TimeSpan.Zero;
    private DateTime _startReference;

    public DispatcherTimerService()
    {
        _timer = new System.Timers.Timer(200);
        _timer.AutoReset = true;
        _timer.Elapsed += HandleElapsed;
    }

    public event EventHandler<TimeSpan>? Tick;

    public bool IsRunning { get; private set; }

    public TimeSpan Elapsed { get; private set; }

    public void Start(TimeSpan offset)
    {
        _offset = offset;
        _startReference = DateTime.Now;
        IsRunning = true;
        _timer.Start();
        RaiseTick();
    }

    public void Pause()
    {
        if (!IsRunning)
        {
            return;
        }

        _timer.Stop();
        _offset = Elapsed;
        IsRunning = false;
        RaiseTick();
    }

    public void Stop()
    {
        _timer.Stop();
        _offset = TimeSpan.Zero;
        Elapsed = TimeSpan.Zero;
        IsRunning = false;
        RaiseTick();
    }

    private void HandleElapsed(object? sender, System.Timers.ElapsedEventArgs e) => RaiseTick();

    private void RaiseTick()
    {
        if (IsRunning)
        {
            Elapsed = _offset + (DateTime.Now - _startReference);
        }
        else
        {
            Elapsed = _offset;
        }

        MainThread.BeginInvokeOnMainThread(() => Tick?.Invoke(this, Elapsed));
    }

    public void Dispose()
    {
        _timer.Elapsed -= HandleElapsed;
        _timer.Stop();
        _timer.Dispose();
    }
}
