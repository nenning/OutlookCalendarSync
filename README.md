# Outlook Calendar Sync Tool

**A .NET console application to sync appointments between all calendars in classic Outlook.**

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Prerequisites](#prerequisites)
4. [Installation](#installation)
5. [Configuration & CLI Usage](#configuration--cli-usage)
   - [Command-Line Options](#command-line-options)
   - [Examples](#examples)
6. [How It Works](#how-it-works)
   - [Architecture](#architecture)
   - [Sync Algorithm](#sync-algorithm)
7. [Troubleshooting](#troubleshooting)
8. [Contributing](#contributing)
9. [License](#license)

---

## Overview

This tool automates n-way blocking of free/busy time between all Office 365 accounts in **classic** Outlook (Windows). It creates "blocker" appointments to prevent double-bookings, cleaning up stale blockers and handling recurring meetings gracefully.

It runs entirely on the desktop (no cloud dependencies), uses only Outlook COM interop, and supports:

- N-way sync between all configured accounts
- Hourly or one-off syncs
- Custom sync windows via start date and duration
- Dry-run mode
- Full reset of blockers
- Heuristics to avoid duplicate or unnecessary blockers

---

## Features

- **N-way sync** between all tenant calendars
- **Date window** specification (`--startdate`, `--days`)
- **Background mode** (`--background`) for hourly running
- **Test mode** (`--test`) to print operations without modifying
- **Reset mode** (`--reset`) to delete all created blockers
- **Skips** user-created events, "block" meetings, and blockers with rooms
- **Recurring handling** with per-occurrence logic
- **Deduplication** of multiple identical occurrences
- **Single-instance enforcement** via named mutex

---

## Prerequisites

- Windows OS with **classic Outlook** (desktop) installed
- .NET 9 SDK
- Visual Studio 2022 or VS Code (optional)
- Office PIAs or COM references available

---

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-org/outlook-calendar-sync.git
   cd outlook-calendar-sync
   ```
2. **Restore and build:**
   ```bash
   dotnet restore
   dotnet build -c Release
   ```

---

## Configuration & CLI Usage

### Command-Line Options

| Flag                 | Type       | Default              | Description                                                        |
| -------------------- | ---------- | -------------------- | ------------------------------------------------------------------ |
| `-b`, `--background` | boolean    | `false`              | Run in background on an STA thread with hourly sync                |
| `-s`, `--startdate`  | `DateTime` | today (`YYYY-MM-DD`) | Start date (inclusive) for sync window                             |
| `-d`, `--days`       | integer    | `60`                 | Number of days into the future to sync                             |
| `-t`, `--test`       | boolean    | `false`              | Test mode: print planned creates/deletes without modifying Outlook |
| `-r`, `--reset`      | boolean    | `false`              | Delete all blockers created by this tool and exit                  |

### Examples

- **One-off sync next 60 days:**

  ```bash
  dotnet run -- -s 2025-04-23 -d 60
  ```

- **Hourly background sync from today:**

  ```bash
  dotnet run -- -b
  ```

- **Dry-run for next 30 days:**

  ```bash
  dotnet run -- -t -d 30
  ```

- **Reset (delete) all blockers, test mode:**

  ```bash
  dotnet run -- -r -t
  ```

---

## How It Works

### Architecture

1. **Entrypoint (Program.cs)** parses CLI options using System.CommandLine.
2. **Mutex** ensures only one instance runs.
3. **STA thread** (optional) hosts Outlook COM calls without freezing the UI.
4. **Sync logic** in `SynchronizeAllAccounts` performs:
   - Date filtering via `Items.Restrict`
   - Exclusion of user meetings and test-only logic
   - Expansion of recurrences with `RecurrencePattern` and exceptions
   - Deduplication of occurrences
   - Creation and deletion of blockers tagged by `BlockerTag` property

### Sync Algorithm

1. **Gather all real meetings** from all accounts into a master list, excluding any appointments already marked as blockers.
2. **For each account (as a target):**
    a. **Identify existing blockers** in the target account's calendar, deduplicating any that share the same unique ID.
    b. **Iterate through the master list of real meetings:**
        i. Skip if the meeting belongs to the current target account.
        ii. Skip if a valid blocker for this meeting already exists.
        iii. Skip all-day, non-busy, or subjects containing "block".
        iv. Skip if an equivalent real meeting already exists in the target calendar (e.g., the user is an attendee on multiple accounts).
        v. Skip meetings with zero duration.
        vi. If none of the above apply, create a new blocker appointment.
    c. **Remove stale blockers:** Any blockers found in step 2a that were not matched with a real meeting in step 2b are deleted.

---

## Troubleshooting

- **COM exceptions**: Ensure Outlook Classic is running and accounts are logged in.
- **Office DLL not found**: Embed interop types or target `net48` if you rely on GAC.
- **Threading deadlocks**: Avoid long CPU-bound work on the STA thread.
- **Missing PIAs**: Install `Microsoft.Office.Interop.Outlook` nuget or add COM reference.

---

## Contributing

1. Fork the repo and create your branch: `git checkout -b feature/YourFeature`
2. Commit your changes: `git commit -m "Add new feature..."`
3. Push to your branch: `git push origin feature/YourFeature`
4. Open a Pull Request.

Please follow the existing coding style, include unit tests for new logic, and update this README with any new instructions.

---

## License

[MIT License](LICENSE.txt)
