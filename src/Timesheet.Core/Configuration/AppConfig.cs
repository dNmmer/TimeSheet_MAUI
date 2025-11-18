namespace Timesheet.Core.Configuration;

/// <summary>
/// Represents persisted application configuration.
/// </summary>
public sealed record AppConfig(string? ExcelPath)
{
    public static AppConfig Default { get; } = new AppConfig(ExcelPath: null);
}
