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
/// Executes n-way sync among all Outlook accounts over the specified window.
/// </summary>
void ExecuteFullSync(DateTime startDate, int daysAhead, bool isTestMode)
{
    Outlook.Application? outlookApplication = null;
    try
    {
        outlookApplication = new Outlook.Application();
        var accounts = outlookApplication.Session.Accounts.Cast<Outlook.Account>().ToList();
        if (accounts.Count < 2)
        {
            Console.WriteLine("Requires at least two accounts in classic Outlook for synchronization.");
            return;
        }

        SynchronizeAllAccounts(accounts, startDate, startDate.AddDays(daysAhead), isTestMode);
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

        foreach (var account in accountList)
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
/// Synchronizes events among a list of Outlook accounts for a given date range.
/// </summary>
void SynchronizeAllAccounts(
    List<Outlook.Account> accounts,
    DateTime startDate,
    DateTime endDate,
    bool isTestMode)
{
    Console.WriteLine($"\nSyncing ({startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}) {(isTestMode ? " [TEST MODE]" : string.Empty)}");

    string dateFilter = $"[Start] >= '{startDate:g}' AND [End] <= '{endDate:g}'";

    // --- Pass 1: Gather all real meetings from all accounts ---
    var allRealMeetings = new Dictionary<(string, DateTime), Outlook.AppointmentItem>();
    var allRealMeetingIdsByAccount = accounts.ToDictionary(acc => acc.DisplayName, _ => new HashSet<string>());

    foreach (var account in accounts)
    {
        var calendar = account.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        var calendarItems = calendar.Items;
        calendarItems.IncludeRecurrences = true;
        calendarItems.Sort("[Start]");
        var appointments = calendarItems.Restrict(dateFilter);

        foreach (Outlook.AppointmentItem appt in appointments)
        {
            if (appt.UserProperties.Find(BlockerTag) == null)
            {
                allRealMeetings[(appt.GlobalAppointmentID, appt.Start)] = appt;
                allRealMeetingIdsByAccount[account.DisplayName].Add(appt.GlobalAppointmentID);
            }
            else
            {
                Marshal.ReleaseComObject(appt); // Release blockers we find along the way
            }
        }
        Marshal.ReleaseComObject(appointments);
        Marshal.ReleaseComObject(calendarItems);
        Marshal.ReleaseComObject(calendar);
    }

    Console.WriteLine($"Found {allRealMeetings.Count} total real meetings across {accounts.Count} accounts.");

    // --- Pass 2: Create missing blockers in each account ---
    foreach (var targetAccount in accounts)
    {
        Console.WriteLine($"\nSyncing blockers for {targetAccount.DisplayName}{(isTestMode ? " [TEST MODE]" : string.Empty)}");
        var targetCalendar = targetAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        var targetItems = targetCalendar.Items;
        targetItems.IncludeRecurrences = true;
        targetItems.Sort("[Start]");
        var targetAppointments = targetItems.Restrict(dateFilter);

        var existingBlockers = targetAppointments
            .Cast<Outlook.AppointmentItem>()
            .Where(a => a.UserProperties.Find(BlockerTag) != null)
            .ToDictionary(a => ((string)a.UserProperties[BlockerTag].Value, a.Start));

        // For every real meeting that exists, ensure a blocker exists in the current target account,
        // unless the meeting belongs to this account.
        foreach (var realMeeting in allRealMeetings.Values)
        {
            string globalId = realMeeting.GlobalAppointmentID;

            // Don't create a blocker in the meeting's own calendar
            if (allRealMeetingIdsByAccount[targetAccount.DisplayName].Contains(globalId))
            {
                continue;
            }

            // If a valid blocker already exists, do nothing.
            if (existingBlockers.ContainsKey((globalId, realMeeting.Start)))
            {
                continue;
            }

            // Skip all-day events, non-busy meetings, or subjects containing "block"
            if (realMeeting.AllDayEvent ||
                realMeeting.BusyStatus != Outlook.OlBusyStatus.olBusy ||
                (realMeeting.Subject ?? string.Empty).IndexOf("block", StringComparison.OrdinalIgnoreCase) >= 0 ||
                realMeeting.Start == realMeeting.End)
            {
                continue;
            }
            
            // Skip if an equivalent real meeting already exists in the target calendar
            // This handles cases where the user is an attendee of the same meeting on multiple accounts.
            if (FindEquivalentMeeting(realMeeting, targetAccount, dateFilter))
            {
                 Console.WriteLine($"  Skipping blocker creation for '{realMeeting.Subject}' at {realMeeting.Start:g} (equivalent meeting exists).");
                 continue;
            }


            // Create the blocker
            if (isTestMode)
            {
                Console.WriteLine($"  [Test] Would create blocker for '{realMeeting.Subject}' at {realMeeting.Start:g}");
            }
            else
            {
                Console.WriteLine($"  Creating blocker for '{realMeeting.Subject}' at {realMeeting.Start:g}");
                var blocker = (Outlook.AppointmentItem)targetItems.Add(Outlook.OlItemType.olAppointmentItem);
                blocker.Subject = "blocker";
                blocker.Start = realMeeting.Start;
                blocker.End = realMeeting.End;
                blocker.AllDayEvent = realMeeting.AllDayEvent;
                blocker.BusyStatus = Outlook.OlBusyStatus.olBusy;
                blocker.ReminderSet = false;
                blocker.UserProperties.Add(BlockerTag, Outlook.OlUserPropertyType.olText).Value = globalId;
                blocker.Save();
                Marshal.ReleaseComObject(blocker);
            }
        }

        // --- Pass 3: Delete stale blockers from the current target account ---
        var realMeetingKeys = new HashSet<(string, DateTime)>(allRealMeetings.Keys);
        foreach (var (key, staleBlocker) in existingBlockers)
        {
            if (!realMeetingKeys.Contains(key))
            {
                if (isTestMode)
                {
                    Console.WriteLine($"  [Test] Would delete stale blocker at {staleBlocker.Start:g}");
                }
                else
                {
                    Console.WriteLine($"  Deleting stale blocker at {staleBlocker.Start:g}");
                    staleBlocker.Delete();
                }
            }
            Marshal.ReleaseComObject(staleBlocker);
        }
        
        Marshal.ReleaseComObject(targetAppointments);
        Marshal.ReleaseComObject(targetItems);
        Marshal.ReleaseComObject(targetCalendar);
    }

    // Release the real meetings gathered in Pass 1
    foreach (var meeting in allRealMeetings.Values)
    {
        Marshal.ReleaseComObject(meeting);
    }
}

/// <summary>
/// Finds if an equivalent real meeting exists in a given account's calendar.
/// </summary>
bool FindEquivalentMeeting(Outlook.AppointmentItem sourceMeeting, Outlook.Account targetAccount, string dateFilter)
{
    var targetCalendar = targetAccount.DeliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
    var targetItems = targetCalendar.Items;
    targetItems.IncludeRecurrences = true;
    targetItems.Sort("[Start]");
    var targetAppointments = targetItems.Restrict(dateFilter);

    string sourceSubject = NormalizeSubject(sourceMeeting.Subject ?? string.Empty);

    bool found = false;
    foreach (Outlook.AppointmentItem targetMeeting in targetAppointments)
    {
        // We only care about real meetings, not blockers
        if (targetMeeting.UserProperties.Find(BlockerTag) != null)
        {
            Marshal.ReleaseComObject(targetMeeting);
            continue;
        }

        if (sourceMeeting.Start == targetMeeting.Start && sourceMeeting.End == targetMeeting.End)
        {
            string targetSubject = NormalizeSubject(targetMeeting.Subject ?? string.Empty);
            int suffixLength = Math.Min(targetSubject.Length, sourceSubject.Length);
            if (suffixLength > 0)
            {
                string sourceSuffix = sourceSubject[^suffixLength..];
                string targetSuffix = targetSubject[^suffixLength..];
                if (string.Equals(sourceSuffix, targetSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;
                }
            }
            else if (string.IsNullOrEmpty(sourceSubject) && string.IsNullOrEmpty(targetSubject))
            {
                found = true; // Both subjects are empty, consider them equivalent
            }
        }
        Marshal.ReleaseComObject(targetMeeting);
        if(found) break;
    }

    Marshal.ReleaseComObject(targetAppointments);
    Marshal.ReleaseComObject(targetItems);
    Marshal.ReleaseComObject(targetCalendar);

    return found;
}