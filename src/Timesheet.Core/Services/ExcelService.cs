using ClosedXML.Excel;
using Timesheet.Core.Models;

namespace Timesheet.Core.Services;

public sealed class ExcelService
{
    public const string ReferenceSheet = "Справочник";
    public const string TimesheetSheet = "Учет времени";
    public const string WorkdaySheet = "Учет рабочего времени";

    public ReferenceData LoadReferenceData(string workbookPath)
    {
        using var workbook = OpenWorkbook(workbookPath);
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == ReferenceSheet)
                    ?? throw new ExcelStructureException($"Workbook must contain sheet '{ReferenceSheet}'.");

        var projects = sheet.RowsUsed()
            .Skip(1)
            .Select(row => row.Cell(1).GetFormattedString().Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .Distinct()
            .ToList();

        var workTypes = sheet.RowsUsed()
            .Skip(1)
            .Select(row => row.Cell(2).GetFormattedString().Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .Distinct()
            .ToList();

        if (projects.Count == 0 || workTypes.Count == 0)
        {
            throw new ExcelStructureException("Лист 'Справочник' должен содержать хотя бы один проект и вид работ.");
        }

        return new ReferenceData(projects, workTypes);
    }

    public void AppendTimeEntry(string workbookPath, TimesheetEntry entry)
    {
        using var workbook = OpenWorkbook(workbookPath);
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == TimesheetSheet)
                    ?? throw new ExcelStructureException($"Workbook must contain sheet '{TimesheetSheet}'.");

        var row = GetFirstEmptyRow(sheet, 2, 4);

        sheet.Cell(row, 1).Value = entry.FinishedAt.Date;
        sheet.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
        sheet.Cell(row, 2).Value = entry.Project;
        sheet.Cell(row, 3).Value = entry.WorkType;
        sheet.Cell(row, 4).Value = entry.Duration;
        sheet.Cell(row, 4).Style.DateFormat.Format = "[h]:mm:ss";

        workbook.Save();
    }

    public WorkdayStartInfo StartWorkday(string workbookPath, DateTime? nowOverride = null)
    {
        using var workbook = OpenWorkbook(workbookPath);
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == WorkdaySheet)
                    ?? workbook.AddWorksheet(WorkdaySheet);

        EnsureWorkdayHeaders(sheet);
        var now = nowOverride ?? DateTime.Now;
        var row = GetFirstEmptyRow(sheet, 2, 4);

        sheet.Cell(row, 1).Value = now.Date;
        sheet.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
        sheet.Cell(row, 2).Value = now.TimeOfDay;
        sheet.Cell(row, 2).Style.DateFormat.Format = "HH:mm";

        workbook.Save();
        return new WorkdayStartInfo(now.ToString("dd.MM.yyyy"), now.ToString("HH:mm"));
    }

    public WorkdayEndInfo EndWorkday(string workbookPath, DateTime? nowOverride = null)
    {
        using var workbook = OpenWorkbook(workbookPath);
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == WorkdaySheet)
                    ?? throw new ExcelStructureException($"Workbook must contain sheet '{WorkdaySheet}'.");

        var row = FindLastOpenWorkdayRow(sheet)
                  ?? throw new ExcelStructureException("Не найдена незавершённая запись рабочего дня.");

        var now = nowOverride ?? DateTime.Now;
        var startDate = GetDate(sheet.Cell(row, 1));
        var startTime = GetTime(sheet.Cell(row, 2));
        var start = startDate + startTime;
        var duration = now - start;
        if (duration < TimeSpan.Zero)
        {
            duration = TimeSpan.Zero;
        }

        sheet.Cell(row, 3).Value = now.TimeOfDay;
        sheet.Cell(row, 3).Style.DateFormat.Format = "HH:mm";
        sheet.Cell(row, 4).Value = duration;
        sheet.Cell(row, 4).Style.DateFormat.Format = "[h]:mm";

        workbook.Save();
        var totalMinutes = (int)Math.Round(duration.TotalMinutes, MidpointRounding.AwayFromZero);
        var hours = totalMinutes / 60;
        var minutes = totalMinutes % 60;
        return new WorkdayEndInfo($"{hours:00}:{minutes:00}");
    }

    public void CreateTemplate(string path)
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet(ReferenceSheet);
        workbook.AddWorksheet(TimesheetSheet);
        workbook.AddWorksheet(WorkdaySheet);

        var reference = workbook.Worksheet(ReferenceSheet);
        reference.Cell(1, 1).Value = "Проект";
        reference.Cell(1, 2).Value = "Тип работ";

        var timesheet = workbook.Worksheet(TimesheetSheet);
        timesheet.Cell(1, 1).Value = "Дата";
        timesheet.Cell(1, 2).Value = "Проект";
        timesheet.Cell(1, 3).Value = "Тип работ";
        timesheet.Cell(1, 4).Value = "Длительность";

        var workday = workbook.Worksheet(WorkdaySheet);
        workday.Cell(1, 1).Value = "Дата";
        workday.Cell(1, 2).Value = "Начало";
        workday.Cell(1, 3).Value = "Окончание";
        workday.Cell(1, 4).Value = "Длительность";

        workbook.SaveAs(path);
    }

    private static XLWorkbook OpenWorkbook(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Excel file not found: {path}", path);
        }

        return new XLWorkbook(path);
    }

    private static int GetFirstEmptyRow(IXLWorksheet sheet, int startRow, int lastColumn)
    {
        var maxRow = sheet.LastRowUsed()?.RowNumber() ?? startRow;
        for (var row = startRow; row <= maxRow; row++)
        {
            var empty = true;
            for (var col = 1; col <= lastColumn; col++)
            {
                if (!sheet.Cell(row, col).IsEmpty())
                {
                    empty = false;
                    break;
                }
            }

            if (empty)
            {
                return row;
            }
        }

        return maxRow + 1;
    }

    private static int? FindLastOpenWorkdayRow(IXLWorksheet sheet)
    {
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = lastRow; row >= 2; row--)
        {
            var hasDate = !sheet.Cell(row, 1).IsEmpty();
            var hasStart = !sheet.Cell(row, 2).IsEmpty();
            var hasEnd = !sheet.Cell(row, 3).IsEmpty();
            if (hasDate && hasStart && !hasEnd)
            {
                return row;
            }
        }

        return null;
    }

    private static void EnsureWorkdayHeaders(IXLWorksheet sheet)
    {
        if (sheet.Cell(1, 1).IsEmpty())
        {
            sheet.Cell(1, 1).Value = "Дата";
            sheet.Cell(1, 2).Value = "Начало";
            sheet.Cell(1, 3).Value = "Окончание";
            sheet.Cell(1, 4).Value = "Длительность";
        }
    }

    private static DateTime GetDate(IXLCell cell)
    {
        if (cell.TryGetValue(out DateTime date))
        {
            return date.Date;
        }

        if (DateTime.TryParse(cell.GetFormattedString(), out date))
        {
            return date.Date;
        }

        throw new ExcelStructureException("Не удалось прочитать дату начала дня.");
    }

    private static TimeSpan GetTime(IXLCell cell)
    {
        if (cell.TryGetValue(out TimeSpan time))
        {
            return time;
        }

        if (cell.TryGetValue(out DateTime dateTime))
        {
            return dateTime.TimeOfDay;
        }

        if (TimeSpan.TryParse(cell.GetFormattedString(), out time))
        {
            return time;
        }

        throw new ExcelStructureException("Не удалось прочитать время начала дня.");
    }
}
