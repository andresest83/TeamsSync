# TeamsSync

Syncs on-premises Active Directory security groups into Microsoft Teams membership, on a
schedule. PowerShell, no dependencies beyond the AD and Teams modules.

**TL;DR:** tag an AD group with its Teams group ID in `extensionAttribute5`, schedule the
script, and Teams membership follows AD. Handles nested groups recursively, refuses to run
if it finds circular nesting, logs daily, and reports failures to a Teams channel.

## Status

Written in 2019 and run in production, when Teams was new and no group writeback path
existed for this. It is kept as a record of that work and is not maintained. It will not
run unmodified today: `Connect-MicrosoftTeams -Credential` relies on basic authentication,
which Microsoft has since disabled, and `Send-MailMessage` is deprecated.

## How it works

A Team is an Office 365 group underneath. The link between the two systems is stored in AD
itself: the Team's Office 365 group ID goes in the AD group's `extensionAttribute5`. No
external mapping file, nothing to keep in sync separately.

Each run:

1. **Pre-flight.** Walks every security group in the target OU looking for circular
   nesting. If any is found, it emails an HTML table of the offenders and aborts. A
   recursive membership walk over a circular graph never terminates, so this check runs
   before anything else.
2. **Discover.** Collects every security group in the OU with `extensionAttribute5` set.
3. **Flatten.** Resolves each AD group to a user list, recursing through nested groups.
   Disabled accounts and admin accounts (UPN prefix `a-`) are excluded.
4. **Reconcile.** Compares that list against the Team's current members and issues the
   `add-teamuser` and `remove-teamuser` calls needed to make Teams match AD.
5. **Report.** Every action goes to a daily log at `C:\Logs\Teams_Sync_<date>.log`. Every
   failure is logged and emailed, and the run continues to the next group rather than
   stopping.

All AD queries target one named domain controller, so a run never acts on partially
replicated membership.

## Prerequisites

Windows with the ActiveDirectory module (RSAT), plus:

```powershell
Install-PackageProvider -Name nuget -MinimumVersion 2.8.5.201 -Force -Scope AllUsers
Install-Module microsoftteams -Scope AllUsers -Force
```

A `C:\Logs` folder, or change `$Logfile`.

## Setup

Variables at the top of `TeamsSync.ps1`:

| Variable | Purpose |
| --- | --- |
| `$server` | Domain controller to query, pinned to avoid replication delays |
| `$myOU` | OU holding the security groups to sync |
| `$sendMailFrom` | Sending account for notifications |
| `$sendMailTo` | Destination. A Teams channel email address works well here |
| `$teamsCredentials` | One account with read/write on Office 365 and Teams, read on AD |

Then set each AD group's `extensionAttribute5` to the matching Team's Office 365 group ID,
and schedule the script with Task Scheduler.

## Origin

Started from a TechNet Gallery script (`Sync-AD-Group-with-Teams-74598786`, offline since
the Gallery was retired in 2020) and grew as production requirements appeared: nested
groups, circular-nesting protection, account filtering, logging and alerting.

## Licence

Apache-2.0. See [LICENSE](LICENSE).
