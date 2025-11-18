namespace TimeSheet_MAUI.Services;

public interface ITimerService : IDisposable
{
    event EventHandler<TimeSpan>? Tick;

    bool IsRunning { get; }

    TimeSpan Elapsed { get; }

    void Start(TimeSpan offset);

    void Pause();

    void Stop();
}
