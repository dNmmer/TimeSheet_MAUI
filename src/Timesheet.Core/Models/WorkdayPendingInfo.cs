using System;

namespace Timesheet.Core.Models;

public sealed record WorkdayPendingInfo(DateTime Date, TimeSpan StartTime);
