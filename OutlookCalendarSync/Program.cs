global using System;
global using System.Collections.Generic;
global using System.Linq;
global using System.Runtime.InteropServices;
global using System.Threading;
global using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using System.CommandLine;

// Define CLI options
var backgroundOption = new Option<bool>(["-b", "--background"],
    description: "Run hourly in a background STA thread");
var daysOption = new Option<int>(["-d", "--days"],
    getDefaultValue: () => 60,
    description: "Number of days into the future to sync");
var testOption = new Option<bool>(["-t", "--test"],
    description: "Test mode: print planned changes without modifying calendars");

var rootCommand = new RootCommand
{
    backgroundOption,
    daysOption,
    testOption
};

rootCommand.SetHandler(async (
    bool runInBackground,
    int daysToSync,
    bool isTestMode) =>
{
    // Ensure only one instance runs
    using var instanceMutex = new Mutex(true, "Global\\OutlookCalendarSyncTool", out bool isNewInstance);
    if (!isNewInstance)
        return;

    if (runInBackground)
    {
        // Start background STA thread for hourly sync
        var syncThread = new Thread(() =>
        {
            using var timer = new PeriodicTimer(TimeSpan.FromHours(1));
            while (timer.WaitForNextTickAsync().AsTask().Result)
                ExecuteSync(daysToSync, isTestMode);
        })
        { IsBackground = true };

        syncThread.SetApartmentState(ApartmentState.STA);
        syncThread.Start();

        // Keep application running
        await Task.Delay(Timeout.Infinite);
    }
    else
    {
        // Single-run sync
        ExecuteSync(daysToSync, isTestMode);
    }
}, backgroundOption, daysOption, testOption);

return await rootCommand.InvokeAsync(args);

/// <summary>
/// Executes a two-way calendar sync between the first two Office 365 accounts in classic Outlook.
/// </summary>
void ExecuteSync(int daysToSync, bool isTestMode)
{
    Outlook.Application? outlookApp = null;
    try
    {
        // Attach to classic Outlook
        outlookApp = new Outlook.Application();
        var accounts = outlookApp.Session.Accounts.Cast<Outlook.Account>().ToList();

        if (accounts.Count < 2)
        {
            Console.WriteLine("At least two accounts must be configured in Outlook.");
            return;
        }

        // Sync calendars both ways
        SyncCalendarsBetweenAccounts(accounts[0], accounts[1], daysToSync, isTestMode);
        SyncCalendarsBetweenAccounts(accounts[1], accounts[0], daysToSync, isTestMode);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Sync failed: {ex}");
    }
    finally
    {
        if (outlookApp is not null)
            Marshal.ReleaseComObject(outlookApp);
    }
}

/// <summary>
/// Syncs calendar events from sourceAccount to targetAccount over the specified time window.
/// </summary>
void SyncCalendarsBetweenAccounts(
    Outlook.Account sourceAccount,
    Outlook.Account targetAccount,
    int daysToSync,
    bool isTestMode)
{
    DateTime now = DateTime.Now;
    DateTime endDate = now.AddDays(daysToSync);
    string filterExpression = $"[Start] >= '{now:g}' AND [Start] <= '{endDate:g}'";

    Console.WriteLine($"\nSyncing from '{sourceAccount.DisplayName}' to '{targetAccount.DisplayName}' " +
                      $"(next {daysToSync} days){(isTestMode ? " [TEST MODE]" : string.Empty)}");

    // Obtain calendars via each account's DeliveryStore
    var sourceCalendar = sourceAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
    var targetCalendar = targetAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

    var sourceItems = sourceCalendar.Items;
    sourceItems.IncludeRecurrences = true;
    var targetItems = targetCalendar.Items;
    targetItems.IncludeRecurrences = true;

    var sourceEvents = sourceItems.Restrict(filterExpression);
    var targetEvents = targetItems.Restrict(filterExpression);

    const string blockerTag = "OutlookCalendarSyncTool";

    // Build list of actual meetings in target calendar (exclude blockers)
    var existingMeetingsInTarget = targetEvents
        .Cast<Outlook.AppointmentItem>()
        .Where(item => item.UserProperties.Find(blockerTag) == null)
        .Select(item => new
        {
            item.Start,
            item.End,
            Organizer = item.Organizer ?? string.Empty,
            Subject = item.Subject ?? string.Empty
        })
        .ToList();

    // Index existing blocker entries by UID and start time
    var existingBlockers = targetEvents
        .Cast<Outlook.AppointmentItem>()
        .Where(item => item.UserProperties.Find(blockerTag) != null)
        .ToDictionary(
            item => $"{item.UserProperties[blockerTag].Value}|{item.Start:o}",
            item => item
        );

    // Iterate source events
    foreach (object rawEvent in sourceEvents)
    {
        if (rawEvent is not Outlook.AppointmentItem appointment)
            continue;

        try
        {
            // Ignore all-day or non-busy appointments
            if (appointment.AllDayEvent || appointment.BusyStatus != Outlook.OlBusyStatus.olBusy)
                continue;

            string subject = appointment.Subject ?? string.Empty;

            // Skip any meeting named exactly "block"
            if (string.Equals(subject, "block", StringComparison.OrdinalIgnoreCase))
                continue;

            string globalId = appointment.GlobalAppointmentID;

            // Expand recurring occurrences and collect unique instances
            var occurrencesList = new List<Outlook.AppointmentItem> { appointment };
            if (appointment.IsRecurring)
            {
                var pattern = appointment.GetRecurrencePattern();
                foreach (Outlook.Exception exceptionEntry in pattern.Exceptions)
                {
                    if (exceptionEntry.Deleted)
                        continue;

                    try
                    {
                        var occurrence = pattern.GetOccurrence(exceptionEntry.OriginalDate);
                        if (occurrence is not null)
                            occurrencesList.Add(occurrence);
                    }
                    catch (COMException)
                    {
                        // Skip if the occurrence no longer exists
                    }
                }
            }

            // Deduplicate by start, end, and subject
            var uniqueOccurrences = occurrencesList
                .GroupBy(evt => new { evt.Start, evt.End, evt.Subject })
                .Select(g => g.First())
                .ToList();

            // Process each unique occurrence
            foreach (var occurrence in uniqueOccurrences)
            {
                if (occurrence.Start < now || occurrence.Start > endDate)
                    continue;

                string key = $"{globalId}|{occurrence.Start:o}";

                // If already blocked, skip
                if (existingBlockers.Remove(key))
                    continue;

                // Skip if a real meeting exists in target
                var words = subject.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                string suffix = (words.Length > 3)
                    ? string.Join(' ', words.Skip(2))
                    : subject;

                bool hasEquivalentInTarget = existingMeetingsInTarget.Any(meeting =>
                    meeting.Start == occurrence.Start &&
                    meeting.End == occurrence.End &&
                    string.Equals(meeting.Organizer, appointment.Organizer, StringComparison.OrdinalIgnoreCase) &&
                    meeting.Subject.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                );
                if (hasEquivalentInTarget)
                    continue;

                // Log or create blocker
                if (isTestMode)
                {
                    Console.WriteLine($"[Test] Would create blocker for '{subject}' at {occurrence.Start}");
                }
                else
                {
                    Console.WriteLine($"Creating blocker for '{subject}' at {occurrence.Start}");
                    var blocker = (Outlook.AppointmentItem)
                        targetCalendar.Items.Add(Outlook.OlItemType.olAppointmentItem);

                    blocker.Subject = $"[Blocked] {subject}";
                    blocker.Start = occurrence.Start;
                    blocker.End = occurrence.End;
                    blocker.AllDayEvent = false;
                    blocker.BusyStatus = Outlook.OlBusyStatus.olBusy;
                    blocker.ReminderSet = false;
                    blocker.UserProperties
                        .Add(blockerTag, Outlook.OlUserPropertyType.olText)
                        .Value = globalId;

                    blocker.Save();
                    Marshal.ReleaseComObject(blocker);
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(appointment);
        }
    }

    // Remove any leftover blockers that no longer match source events
    foreach (var staleBlocker in existingBlockers.Values)
    {
        if (isTestMode)
            Console.WriteLine($"[Test] Would delete stale blocker at {staleBlocker.Start}");
        else
            Console.WriteLine($"Deleting stale blocker at {staleBlocker.Start}");

        if (!isTestMode)
            staleBlocker.Delete();

        Marshal.ReleaseComObject(staleBlocker);
    }

    // Cleanup COM objects
    Marshal.ReleaseComObject(sourceEvents);
    Marshal.ReleaseComObject(sourceItems);
    Marshal.ReleaseComObject(targetEvents);
    Marshal.ReleaseComObject(targetItems);
}