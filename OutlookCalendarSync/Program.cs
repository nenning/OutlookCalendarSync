global using System;
global using System.Collections.Generic;
global using System.Linq;
global using System.Runtime.InteropServices;
global using System.Threading;
global using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using System.CommandLine;
using System.CommandLine.Invocation;

// Tag used to mark blocker appointments
const string BlockerTag = "MySyncToolId";

/// <summary>
/// Removes any prefix before the first ':' and any bracketed tag or word: at the start.
/// </summary>
static string NormalizeSubject(string rawSubject)
{
    if (string.IsNullOrWhiteSpace(rawSubject))
        return string.Empty;
    int colonIndex = rawSubject.IndexOf(':');
    string afterColon = colonIndex >= 0
        ? rawSubject[(colonIndex + 1)..]
        : rawSubject;
    // Strip any leading [TAG] or word: prefix
    return System.Text.RegularExpressions.Regex.Replace(
        afterColon.Trim(),
        "^(?:\\[[^\\]]*\\]|\\w+:)\\s*",
        string.Empty,
        System.Text.RegularExpressions.RegexOptions.IgnoreCase
    ).Trim();
}

// Define CLI options
var backgroundOption = new Option<bool>(new[] { "-b", "--background" },
    description: "Run hourly in a background STA thread.");
var daysAheadOption = new Option<int>(new[] { "-d", "--days" },
    getDefaultValue: () => 60,
    description: "Number of days into the future to sync.");
var startDateOption = new Option<DateTime>(new[] { "-s", "--startdate" },
    getDefaultValue: () => DateTime.Now.Date,
    description: "Start date (inclusive) in format yyyy-MM-dd.");
var testModeOption = new Option<bool>(new[] { "-t", "--test" },
    description: "Test mode: print planned changes without modifying calendars.");
var resetModeOption = new Option<bool>(new[] { "-r", "--reset" },
    description: "Delete all blockers created by this tool and exit.");

// Assemble root command
var rootCommand = new RootCommand
{
    backgroundOption,
    daysAheadOption,
    startDateOption,
    testModeOption,
    resetModeOption
};

rootCommand.SetHandler(async (
    bool runInBackground,
    int daysAhead,
    DateTime startDate,
    bool isTestMode,
    bool isResetMode) =>
{
    // Ensure single instance
    using var instanceMutex = new Mutex(true, "Global\\CalendarSyncTool", out bool isFirstInstance);
    if (!isFirstInstance)
        return;

    if (isResetMode)
    {
        DeleteAllBlockers(isTestMode);
        return;
    }

    if (runInBackground)
    {
        // Background STA thread for hourly sync
        var syncThread = new Thread(() =>
        {
            using var hourlyTimer = new PeriodicTimer(TimeSpan.FromHours(1));
            while (hourlyTimer.WaitForNextTickAsync().AsTask().Result)
                ExecuteFullSync(startDate, daysAhead, isTestMode);
        })
        { IsBackground = true };

        syncThread.SetApartmentState(ApartmentState.STA);
        syncThread.Start();
        await Task.Delay(Timeout.Infinite);
    }
    else
    {
        ExecuteFullSync(startDate, daysAhead, isTestMode);
    }
}, backgroundOption, daysAheadOption, startDateOption, testModeOption, resetModeOption);

return await rootCommand.InvokeAsync(args);

/// <summary>
/// Executes two-way sync between the first two Outlook accounts over the specified window.
/// </summary>
void ExecuteFullSync(DateTime startDate, int daysAhead, bool isTestMode)
{
    DateTime endDate = startDate.AddDays(daysAhead);
    Outlook.Application? outlookApplication = null;
    try
    {
        outlookApplication = new Outlook.Application();
        var accountList = outlookApplication.Session.Accounts.Cast<Outlook.Account>().ToList();
        if (accountList.Count < 2)
        {
            Console.WriteLine("Requires two accounts in classic Outlook.");
            return;
        }

        // Sync calendars in both directions
        SynchronizeBetweenAccounts(accountList[0], accountList[1], startDate, endDate, isTestMode);
        SynchronizeBetweenAccounts(accountList[1], accountList[0], startDate, endDate, isTestMode);
    }
    catch (System.Exception exception)
    {
        Console.WriteLine($"Sync failed: {exception.Message}");
    }
    finally
    {
        if (outlookApplication is not null)
            Marshal.ReleaseComObject(outlookApplication);
    }
}

/// <summary>
/// Deletes all blocker appointments created by this tool, skipping those with a meeting room.
/// </summary>
void DeleteAllBlockers(bool isTestMode)
{
    Outlook.Application? outlookApplication = null;
    try
    {
        outlookApplication = new Outlook.Application();
        var accountList = outlookApplication.Session.Accounts.Cast<Outlook.Account>().ToList();
        if (accountList.Count < 2)
            return;

        foreach (var account in accountList.Take(2))
        {
            var calendarFolder = account.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            var calendarItems = calendarFolder.Items;
            calendarItems.IncludeRecurrences = false;

            var blockerAppointments = calendarItems
                .Cast<Outlook.AppointmentItem>()
                .Where(appointment => appointment.UserProperties.Find(BlockerTag) != null)
                .ToList();

            Console.WriteLine($"\nResetting blockers for {account.DisplayName}:");
            foreach (var blockerAppointment in blockerAppointments)
            {
                bool hasRoom = !string.IsNullOrEmpty(blockerAppointment.Location);
                if (hasRoom)
                {
                    Console.WriteLine($" Skipping {blockerAppointment.Start} (meeting room assigned)");
                }
                else if (isTestMode)
                {
                    Console.WriteLine($" [Test] Would delete blocker at {blockerAppointment.Start}");
                }
                else
                {
                    Console.WriteLine($" Deleting blocker at {blockerAppointment.Start}");
                    blockerAppointment.Delete();
                }
                Marshal.ReleaseComObject(blockerAppointment);
            }

            Marshal.ReleaseComObject(calendarItems);
        }
    }
    catch (System.Exception exception)
    {
        Console.WriteLine($"Reset failed: {exception.Message}");
    }
    finally
    {
        if (outlookApplication is not null)
            Marshal.ReleaseComObject(outlookApplication);
    }
}

/// <summary>
/// Synchronizes events from a source account to a target account within a date range.
/// </summary>
void SynchronizeBetweenAccounts(
    Outlook.Account sourceAccount,
    Outlook.Account targetAccount,
    DateTime startDate,
    DateTime endDate,
    bool isTestMode)
{
    string dateFilter = $"[Start] >= '{startDate:g}' AND [Start] <= '{endDate:g}'";
    Console.WriteLine($"\nSync {sourceAccount.DisplayName} -> {targetAccount.DisplayName} " +
                      $"({startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd})" +
                      $"{(isTestMode ? " [TEST MODE]" : string.Empty)}");

    var sourceCalendarFolder = sourceAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
    var targetCalendarFolder = targetAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

    var sourceCalendarItems = sourceCalendarFolder.Items;
    sourceCalendarItems.Sort("[Start]");
    sourceCalendarItems.IncludeRecurrences = true;
    var targetCalendarItems = targetCalendarFolder.Items;
    targetCalendarItems.Sort("[Start]");
    targetCalendarItems.IncludeRecurrences = true;

    var sourceEventCollection = sourceCalendarItems.Restrict(dateFilter);
    var targetEventCollection = targetCalendarItems.Restrict(dateFilter);

    // Build list of real meetings in target calendar\    
    var realMeetingsInTargetList = targetEventCollection
        .Cast<Outlook.AppointmentItem>()
        .Where(appointment => appointment.UserProperties.Find(BlockerTag) == null)
        .Select(appointment => new
        {
            appointment.Start,
            appointment.End,
            appointment.Subject
        })
        .ToList();

    // Map existing blockers by unique key
    var existingBlockersMap = targetEventCollection
        .Cast<Outlook.AppointmentItem>()
        .Where(appointment => appointment.UserProperties.Find(BlockerTag) != null)
        .ToDictionary(
            appointment => $"{appointment.UserProperties[BlockerTag].Value}|{appointment.Start:o}",
            appointment => appointment
        );

    foreach (object rawEvent in sourceEventCollection)
    {
        if (rawEvent is not Outlook.AppointmentItem appointmentItem)
            continue;

        try
        {
            // Skip if already blocked or not busy/all-day
            if (appointmentItem.UserProperties.Find(BlockerTag) != null)
                continue;
            if (appointmentItem.AllDayEvent || appointmentItem.BusyStatus != Outlook.OlBusyStatus.olBusy)
                continue;

            string rawSubject = appointmentItem.Subject ?? string.Empty;
            string normalizedSubject = NormalizeSubject(rawSubject);

            // Skip meetings titled exactly "block" or "blocker"
            if (string.Equals(normalizedSubject, "block", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalizedSubject, "blocker", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            DateTime occurrenceStart = appointmentItem.Start;
            DateTime occurrenceEnd = appointmentItem.End;

            // Skip zero-duration
            if (occurrenceEnd <= occurrenceStart)
                continue;

            // Skip outside window
            if (occurrenceStart < startDate || occurrenceStart > endDate)
                continue;

            string uniqueKey = $"{appointmentItem.GlobalAppointmentID}|{occurrenceStart:o}";
            if (existingBlockersMap.Remove(uniqueKey))
                continue;

            // Check for equivalent real meeting by suffix comparison
            bool equivalentMeetingExists = realMeetingsInTargetList.Any(meeting =>
            {
                if (meeting.Start != occurrenceStart || meeting.End != occurrenceEnd)
                    return false;
                string targetSubject = NormalizeSubject(meeting.Subject);
                int suffixLength = Math.Min(targetSubject.Length, normalizedSubject.Length);
                string sourceSuffix = normalizedSubject[^suffixLength..];
                string targetSuffix = targetSubject[^suffixLength..];
                return string.Equals(sourceSuffix, targetSuffix, StringComparison.OrdinalIgnoreCase);
            });

            if (equivalentMeetingExists)
                continue;

            // Create blocker or simulate
            if (isTestMode)
            {
                Console.WriteLine($"[Test] Would create blocker at {occurrenceStart} ('{rawSubject}')");
            }
            else
            {
                Console.WriteLine($"Creating blocker at {occurrenceStart} ('{rawSubject}')");
                var blockerAppointment = (Outlook.AppointmentItem)
                    targetCalendarFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);
                blockerAppointment.Subject = "blocker";
                blockerAppointment.Start = occurrenceStart;
                blockerAppointment.End = occurrenceEnd;
                blockerAppointment.AllDayEvent = false;
                blockerAppointment.BusyStatus = Outlook.OlBusyStatus.olBusy;
                blockerAppointment.ReminderSet = false;
                blockerAppointment.UserProperties
                    .Add(BlockerTag, Outlook.OlUserPropertyType.olText)
                    .Value = appointmentItem.GlobalAppointmentID;
                blockerAppointment.Save();
                Marshal.ReleaseComObject(blockerAppointment);
            }
        }
        finally
        {
            Marshal.ReleaseComObject(appointmentItem);
        }
    }

    // Delete stale blockers
    foreach (var staleBlocker in existingBlockersMap.Values)
    {
        if (isTestMode)
            Console.WriteLine($"[Test] Would delete stale blocker at {staleBlocker.Start}");
        else
        {
            Console.WriteLine($"Deleting stale blocker at {staleBlocker.Start}");
            staleBlocker.Delete();
        }
        Marshal.ReleaseComObject(staleBlocker);
    }

    // Cleanup COM objects
    Marshal.ReleaseComObject(sourceEventCollection);
    Marshal.ReleaseComObject(sourceCalendarItems);
    Marshal.ReleaseComObject(targetEventCollection);
    Marshal.ReleaseComObject(targetCalendarItems);
}