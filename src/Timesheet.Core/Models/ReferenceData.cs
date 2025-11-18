namespace Timesheet.Core.Models;

public sealed record ReferenceData(IReadOnlyList<string> Projects, IReadOnlyList<string> WorkTypes);
