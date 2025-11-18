namespace Timesheet.Core.Models;

public sealed record TimesheetEntry(string Project, string WorkType, TimeSpan Duration, DateTime FinishedAt);
