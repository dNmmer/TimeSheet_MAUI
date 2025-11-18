namespace Timesheet.Core.Services;

public sealed class ExcelStructureException : Exception
{
    public ExcelStructureException(string message) : base(message)
    {
    }
}
