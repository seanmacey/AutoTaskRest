# AutoTaskRest

A PowerShell module for interacting with the Autotask PSA REST API. Provides functions for reading and updating companies, tickets, time entries, contacts, and engineer assignments, along with calendar synchronisation to Microsoft 365 and Outlook.

---

## Table of contents

- [Requirements](#requirements)
- [Installation](#installation)
- [First-time setup](#first-time-setup)
- [Primary and secondary engineer assignments](#primary-and-secondary-engineer-assignments)
- [Reporting — time entries, tickets, and daily stats](#reporting--time-entries-tickets-and-daily-stats)
- [Company management](#company-management)
- [Calendar synchronisation](#calendar-synchronisation)
- [Function reference](#function-reference)
- [Configuration constants](#configuration-constants)
- [How to get Help](#how-to-get-help--man-or-get-help)
- [Notes about AuTotask Integration](#notes-about-the-integration-of-autotask-to-this-script)

---

## Requirements

- PowerShell 7+  (it does not run on PowerShell 5.1)
  - so run Pwsh not powershell 
- An Autotask API user account with an API Integration Code
- For calendar sync: the [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation) (`Install-Module Microsoft.Graph`)

---

## Updating the CodeSIgning validation

If you have modified the AutotaskRest script then...  you need to reapply the public key into it. to do this see the Certificates-forAutoTask\Certificates.mz file for instructions on how to do that. - But short answer is run the below command - after downloading the public certificate

```powershell
copy-item .\AutoTaskRest.ps1 .\AutoTaskRest.psm1 -force
$CerttoApply = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object Subject -eq "CN=Sean Macey"

#Get-ChildItem  -Recurse -Include *.ps1,*.psm1,*.psd1 |
Get-ChildItem   *.ps1,*.psm1,*.psd1 |
  ForEach-Object {
    Set-AuthenticodeSignature $_ $CerttoApply
  }


```

### Install Self Signed publicCertificate

Installing the Self Signed Certificate that is stored along with this module, will allow the module to be run under the 'remote-signed' setting, and will allow you to identify if the Module has been modified after the time it was signed.

- run this to install the public certiicate on a machine that will run the autotaskrest module
- If installing AutoTaskRest via **install-autotaskRest** you do not need to add this certificate - the installer will do this for you (if needed )


if yhou want to check if this certificate is installed to both locations then
```
Get-ChildItem Cert:CurrentUser\TrustedPublisher |
Where-Object { $_.Subject -like "*Sean Macey*" }
Get-ChildItem Cert:\LocalMachine\Root |
Where-Object { $_.Subject -like "*Sean Macey*" }

```




### Option 1 — Install as a PowerShell module (recommended)

Installing as a module means all functions are available in every PowerShell session automatically without needing to dot-source the file.
Download the install-autotaskrest.ps1 script

- either save it locally and run it via an pswh administrative session
- or copy the text from within the script and run that directly in a pwsh admin window
- If you run the install script from a powershell 5 session it will prompt you to install pwsh 7, or to use pwsh instead provided you already have it installed. 

<https://rmm.kissit.co.nz/webdocs/AutoTaskRest/install-autotaskrest.ps1>



**Verify the installation**

Open a new PowerShell window and run:

```powershell
Get-Module -ListAvailable AutoTaskRest
Get-Command -Module AutoTaskRest
```

You should see all 41 functions listed.

### Option 2 — Dot-source per session

If you prefer not to install as a module, dot-source the file at the start of each session or script:

```powershell
. "C:\path\to\AutoTaskRest.ps1"
```

---

## Uninustalling the module

download and run <https://rmm.kissit.co.nz/webdocs/AutoTaskRest/remove-autotaskrest.ps1>

---

## First-time setup

Before any function can connect to Autotask, you need to save your API credentials. This only needs to be done once per machine. Credentials are encrypted using Windows DPAPI and stored in `$HOME\kiss-atapi\kissAtapilogin.json` — they are tied to your Windows user account and cannot be read by other users.

**You will need:**
- Your Autotask **API username** (usually an email address)
- Your Autotask **API user password**
- Your **API Integration Code** (found in Autotask under Admin → API Integration Codes)

**Run the setup wizard:**

```powershell
Set-ATLogin
```

The wizard will prompt for each value interactively. The password and API Integration Code are entered as secure strings and are never displayed on screen. When complete, it tests the connection and saves the credentials only if the test succeeds.

**Test that your saved credentials work:**

```powershell
Test-ATConnection
```

**Set your engineer ID** (used by `-ForMeOnly` switches in reporting functions):

```powershell
# Find your engineer record first
Get-ATEngineers | Where-Object email -like "*yourname*" | Select-Object id, userName, email

# Then save your ID
Set-ATEngineerToReport -id 29683001

# Or save by email
Set-ATEngineerToReport -email "your.name@company.co.nz"

# Verify it was saved
Get-ATEngineerToReport
```

---

## Primary and secondary engineer assignments

Engineer assignments are stored as alert text on each company record in Autotask. The module reads and writes these alerts to track which engineer is primary and which is secondary for each customer.

### View current assignments

**Get all company assignments:**

```powershell
$assignments = Get-ATCompanyEngineers
$assignments | Select-Object CompanyID, Primary, Secondary | Sort-Object Primary
```

**Include company detail (name, branch, active status):**

```powershell
$assignments = Get-ATCompanyEngineers -IncludeCompanyDetail
$assignments | Select-Object Company, Primary, Secondary, Branch, isActive
```

**Export to CSV for review:**

```powershell
Get-ATCompanyEngineers -IncludeCompanyDetail |
    Export-Csv .\EngineerAssignments.csv -NoTypeInformation
```

**Find all companies assigned to a specific engineer:**

```powershell
Get-ATCompanyEngineers -IncludeCompanyDetail |
    Where-Object Primary -eq "Sean" |
    Select-Object Company, Secondary, Branch
```

**Flag inactive companies that still have assignments:**

```powershell
Get-ATCompanyEngineers -IncludeCompanyDetail |
    Where-Object { -not $_.isActive -and ($_.Primary -or $_.Secondary) }
```

### Update assignments

**Assign engineers to a single company by name:**

```powershell
Set-ATCompanyEngineers -CompanyName "Matamata Medical Center" -Primary "Sean" -Secondary "Antony"
```

**Assign by company ID:**

```powershell
Set-ATCompanyEngineers -ID 29762990 -Primary "Sean" -Secondary "Antony"
```

**Remove a secondary assignment (leave primary unchanged):**

```powershell
Set-ATCompanyEngineers -CompanyName "Acme Ltd" -Primary "Sean" -Secondary "null"
```

**Remove both assignments:**

```powershell
Set-ATCompanyEngineers -CompanyName "Acme Ltd" -Primary "null" -Secondary "null"
```

**Bulk update from a CSV file:**

Create a CSV with columns `CompanyName` (or `CompanyID`), `Primary`, and `Secondary`:

```
CompanyName,Primary,Secondary
Acme Ltd,Sean,Antony
Matamata Medical,Jane,
Old Customer Ltd,null,null
```

Then import and pipe it:

```powershell
Import-Csv .\PrimaryEngineers.csv | Set-ATCompanyEngineers
```

**Remove assignments for all inactive companies:**

```powershell
Get-ATCompanyEngineers -IncludeCompanyDetail -UnassignInactiveCustomers |
    Where-Object { -not $_.isActive } |
    ForEach-Object {
        Set-ATCompanyEngineers -ID $_.CompanyID -Primary "null" -Secondary "null"
    }
```

---

## Reporting — time entries, tickets, and daily stats

### Weekly timesheet summary

Get a quick billable/non-billable breakdown for the current week:

```powershell
# Your own hours for the current week
Get-ATWeeklySummary

# Last 2 weeks, all engineers
Get-ATWeeklySummary -lastXWeeks 2 -AllEngineers
```

The summary prints to the console and also returns a `PSCustomObject` array with the properties `Engineer`, `hoursBillable`, `hoursNonBillable`, `hoursInternal`, `weeks`, `startDate`, and `endDate`.

### Detailed time entries

```powershell
# Last 3 months (default)
$entries = Get-ATTimeEntries

# Last 4 weeks
$entries = Get-ATTimeEntries -LastXWeeks 4

# Last 6 months with all enrichment fields
$entries = Get-ATTimeEntries -LastxMonths 6 `
    -IncludeBillingDetails `
    -includeTicketDetails `
    -includeEngineerDetails

# Just your own entries
$entries = Get-ATTimeEntries -LastXWeeks 2 -ForMeOnly

# For a specific engineer
$entries = Get-ATTimeEntries -LastXWeeks 4 -ForResouerceID 29683001

# For a specific ticket
$entries = Get-ATTimeEntries -ForTicketNumber "T20260101.0001"
```

**Useful things to do with time entries:**

```powershell
# Total billable hours per engineer
$entries | Group-Object Engineer |
    Select-Object Name, @{n='BillableHrs'; e={ ($_.Group | Measure-Object hoursBillable -Sum).Sum }}

# Non-billable entries only
$entries | Where-Object isNonBillable -eq $true

# After-hours entries
$entries | Where-Object kissWorkType -like "*AfterHrs*"

# Entries for a specific company
$entries | Where-Object Company -eq "Acme Ltd"

# Export to CSV
$entries | Export-Csv .\TimeEntries.csv -NoTypeInformation
```

### Daily time statistics

`Get-ATDailyTimeStats` builds a per-engineer, per-day breakdown. It compares actual hours worked against each engineer's expected daily hours (pulled from Autotask ResourceDailyAvailabilities).

```powershell
$entries = Get-ATTimeEntries -LastXWeeks 4 `
    -IncludeBillingDetails -includeTicketDetails -includeEngineerDetails

$daily = Get-ATDailyTimeStats -TimeEntries $entries
$daily | Select-Object Resource, workDate, hoursWorked, HoursExpectedPerDay, HrsClient |
    Sort-Object Resource, workDate
```

### Ticket reports

```powershell
# All open (non-completed) tickets with an assigned engineer
$open = Get-ATTickets -IncludeAllNonComplete -GetCompanyNames

# Tickets active in the last 30 days
$recent = Get-ATTickets -LastActionFromDate (Get-Date).AddDays(-30) -GetCompanyNames

# Tickets for specific companies
$tickets = Get-ATTickets -CompanyIDs @(29762985, 29740186)

# Tickets by ID
$tickets = Get-ATTickets -ids @(12345, 67890)

# Find all RMM tickets
$rmm = Get-ATTickets -TitleBeginsWith "RMM" -LastActionFromDate (Get-Date).AddMonths(-3)

# Tickets containing a keyword in the title
$search = Get-ATTickets -TitleContains "firewall" -LastActionFromDate (Get-Date).AddMonths(-1)

# Completed tickets in the last 2 weeks
$done = Get-ATTickets -ExcludeNonComplete -LastxWeeks 2 -GetCompanyNames
```

**Useful ticket queries:**

```powershell
# Count open tickets per company
$open | Group-Object CompanyName |
    Select-Object Name, Count | Sort-Object Count -Descending

# Tickets not yet assigned
$open | Where-Object { -not $_.assignedResourceID }

# Recently completed, by queue
$done | Group-Object QueueName |
    Select-Object Name, Count
```

### Bulk CSV/JSON exports

These functions create ready-to-use data files suitable for Power BI, Excel, or archiving.

**Export everything (time entries, tickets, daily stats, billing codes, engineers):**

```powershell
# CSV to current directory
Export-ATTimeRecords -LastxMonths 3

# CSV to a specific path
Export-ATTimeRecords -LastxMonths 6 -path "W:\Autotask\Reports"

# JSON format
Export-ATTimeRecords -LastxMonths 3 -exportType JSON
```

This creates the following files:

| File | Contents |
|---|---|
| `KissTimeEntries.csv` | All time entries with enrichment fields |
| `KissDaily.csv` | Per-engineer per-day summary |
| `KissTickets.csv` | All tickets active in the period |
| `KissBillingCodes.csv` | Billing code lookup table |
| `KissEngineers.csv` | Engineer records |

**Export company list:**

```powershell
Export-ATCompanies
Export-ATCompanies -path "W:\Autotask\Reports"
```

Creates `KissAtCompanies.csv` and `KissAtClassificationIcons.csv`.

**Export ticket CSVs:**

```powershell
# All open tickets only
Export-ATTickets

# Open tickets + tickets active in the last 3 months
Export-ATTickets -WhereLastActionOccurWithinLastMonths 3 -path "W:\Reports"
```

---

## Company management

### Look up companies

```powershell
# All active companies
$companies = Get-ATCompanies

# Search by name (partial match)
Get-ATCompanies -CompanyName "Medical"

# Exact name match
Get-ATCompanies -CompanyName "Matamata Medical Center" -exactNameMatch

# By ID
Get-ATCompanies -id 29762985

# Multiple IDs (chunked automatically)
Get-ATCompanies -id @(29762985, 29740186, 29761818)

# Include inactive companies
Get-ATCompanies -includeInactive

# Include primary/secondary engineer info
Get-ATCompanies -GetEngineers
```

### Update company fields

```powershell
# Change branch
Set-ATCompanies -CompanyID 29762985 -branch "Tauranga"

# Change manager and classification
Set-ATCompanies -CompanyName "Kiss IT" -Manager "Sean Macey" -Classification "Residential"

# Deactivate a company
Set-ATCompanies -CompanyID 29762985 -isActive $false

# Bulk update from CSV
Import-Csv .\CompanyUpdates.csv | Set-ATCompanies
```

### Contacts

```powershell
# All contacts
Get-ATContacts

# By ID
Get-ATContacts -ID 29692052

# By email
Get-ATContacts -eMail "bob@example.com"

# Update a contact's email or bulk-email opt-out status
$contact = Get-ATContacts -eMail "old@example.com"
$contact | Set-ATContact -isOptedOutFromBulkEmail True

# Set email to unknown (for GDPR cleanup)
$contact | Set-ATContact -SetunknownEmail
```

---

## Calendar synchronisation

The module can sync Autotask time entries to either Microsoft 365 (via Microsoft Graph) or a local Outlook installation (via COM automation). Each time entry becomes a calendar event; subsequent runs update existing events in place rather than creating duplicates.

### Microsoft 365 calendar

Requires the Microsoft Graph PowerShell SDK and a `Calendars.ReadWrite` consent grant.

```powershell
# Install the Graph SDK if not already present
Install-Module Microsoft.Graph -Scope CurrentUser

# Sync the current week
Sync-AT365Calendar

# Sync the last 2 weeks, create new events with full ticket detail
Sync-AT365Calendar -LastXWeeks 2

# Sync without fetching ticket titles (faster, less detail on new events)
Sync-AT365Calendar -LastXWeeks 2 -DoNotProvideDetaledInfoOnNew

# Sync and also update existing events with full detail
Sync-AT365Calendar -LastXWeeks 2 -UpdateExistingInDetail

# Just read your 365 calendar events (no writes)
Get-AT365CalendarEvents -LastXWeeks 4
Get-AT365CalendarEvents -OnlyGetItemsWithAutotaskTimeEntryIDsInBody
```

### Outlook desktop calendar

Requires Outlook installed and configured on the local machine.

```powershell
# Sync the current week
Sync-ATOutlookCalendar

# Sync 2 weeks and update existing events with full ticket detail
Sync-ATOutlookCalendar -LastXWeeks 2 -UpdateExistingInDetail
```

---

## Function reference

### Authentication and setup

| Function | Description |
|---|---|
| `Set-ATLogin` | Interactive wizard to save API credentials (encrypted via DPAPI) |
| `Test-ATConnection` | Verifies saved credentials against the live API |
| `Get-ATEngineerToReport` | Returns the saved engineer ID used by `-ForMeOnly` switches |
| `Set-ATEngineerToReport` | Saves an engineer ID or email as the default reporting engineer |

### Companies

| Function | Description |
|---|---|
| `Get-ATCompanies` | Returns active or all companies, with optional name/ID filtering |
| `Set-ATCompanies` | Updates branch, manager, classification, or active status |
| `Get-ATClassificationIcons` | Returns the classification icon lookup table |
| `Get-ATMostRecentCompanyTicket` | Returns the most recently completed ticket per company |
| `Get-ATLastQuickNote` | Returns the most recent quick note per company |

### Contacts

| Function | Description |
|---|---|
| `Get-ATContacts` | Returns contacts filtered by ID or email |
| `Set-ATContact` | Updates a contact's email address or bulk-email opt-out status |

### Engineer assignments

| Function | Description |
|---|---|
| `Get-ATCompanyEngineers` | Returns primary/secondary engineer assignments for all companies |
| `Set-ATCompanyEngineers` | Sets or clears primary/secondary assignments for one or more companies |
| `Get-ATEngineers` | Returns Autotask resource (engineer) records |
| `Get-ATCompanyAlert` | Returns alert text for a specific alert type on a company |
| `Get-ATCompanyChildAlerts` | Returns all alerts for a company |

### Time entries

| Function | Description |
|---|---|
| `Get-ATTimeEntries` | Polls time entries with date, resource, ticket, and billing filters |
| `Get-ATWeeklySummary` | Prints and returns a per-engineer weekly hours summary |
| `Get-ATDailyTimeStats` | Builds a per-engineer per-day breakdown from a time-entry array |
| `Set-ATInternalTicketTime` | Annotates internal-company time entries with classification fields |

### Tickets

| Function | Description |
|---|---|
| `Get-ATTickets` | Returns tickets with flexible filtering, chunked for large ID sets |
| `Get-ATTicketFieldInfo` | Returns picklist values for ticket queues, statuses, and categories |
| `Find-ATCompaniesInTickets` | Returns the company record for each company referenced in a ticket list |

### Billing and roles

| Function | Description |
|---|---|
| `Get-ATBillingCodes` | Returns all billing code records |
| `Get-ATRoles` | Returns all role records |

### Exports

| Function | Description |
|---|---|
| `Export-ATTimeRecords` | Exports time entries, daily stats, tickets, billing codes, and engineers to CSV or JSON |
| `Export-ATCompanies` | Exports companies and classification icons to CSV |
| `Export-ATTickets` | Exports open and/or recently-actioned tickets to CSV |

### Calendar sync

| Function | Description |
|---|---|
| `Sync-AT365Calendar` | Syncs time entries to Microsoft 365 calendar via Microsoft Graph |
| `Sync-ATOutlookCalendar` | Syncs time entries to local Outlook calendar via COM automation |
| `Get-AT365CalendarEvents` | Returns 365 calendar events for a date range |
| `New-AT365CalendarEvent` | Creates a new 365 calendar event |
| `Update-AT365CalendarEvent` | Updates an existing 365 calendar event |

### Utilities

| Function | Description |
|---|---|
| `Get-ATWeekStart` | Returns the date of the Sunday N weeks ago at midnight |
| `Test-ATWorkingDay` | Returns `$true` if a date falls on a weekday |
| `Convert-ObjArrayDateTimesToSearchableStrings` | Converts DateTime properties to sortable ISO 8601 strings |
| `convertto-escapedString` | Percent-encodes `&` in strings for use in API query filters |

### Internal / low-level

| Function | Description |
|---|---|
| `Get-ATCredentialHeader` | Builds the authentication header from saved credentials (private) |
| `Invoke-ATQuery` | Paginated GET wrapper for all read operations |
| `Invoke-ATREST` | Single-request wrapper for write operations (PUT, POST, PATCH, DELETE) |

---

## Configuration constants

These script-level constants are set at the top of `AutoTaskRest.ps1`. Edit them if your Autotask environment uses different IDs.

| Variable | Default | Purpose |
|---|---|---|
| `$script:ATnonBillableCodes` | `@(29682861)` | Billing code IDs treated as non-billable |
| `$script:ATInternalClasificationCode` | `200` | Company classification ID for internal/Kiss IT companies |
| `$script:ATtaskCodesRMM` | `29712660` | Billing code ID for RMM time entries |

Credentials and login state are stored at:

```
$HOME\kiss-atapi\kissAtapilogin.json
```

This file is created automatically by `Set-ATLogin` and is encrypted with the Windows DPAPI tied to your user account.


## how to Get Help  (MAN or Get-Help)
Most of the functions have inbuilt help- - just euse the **Man** or **Get-Help** commands (example below)

``` powershell
man Get-ATCompanies
  NAME
  Get-ATCompanies
  SYNOPSIS
  returns a list of companies (or just one of)
```

## Notes about the integration of AutoTask to this script

I gleaned information from <https://autotask.net/help/DeveloperHelp/Content/APIs/REST/REST_API_Home.htm> to build these scripts

## How datetime fields are handled

the API needs to be date local invariant, so the searchable date text date format is used
EXAMPLE  When making a ContractServiceAdjustments call, the effective date is submitted as **2023-10-09T02:00:00.00**, that is, 2 AM on October 9. Because the API intakes call in UTC, if that call is made to a US database (UTC + 5), it would seem to change the effective date to October 8th at 9 PM, due to the time zone conversion.
However, because there is no time field in the UI for service adjustments, we don't convert timezone datetime values for date-only fields, we just set the time portion to midnight and accept the date value.

In the example above, the datetime would be saved in the database as **2023-10-08T00:00:00.00**.
powershell can create this format  example: ```$Monthstart.ToString("yyyy-MM-ddTHH:mm:ss")```

## Filter operators

Most calls to the API will need one or more filter operators to indicate the type of query you'd like the API to perform. The table below lists the available operators and their definitions.
You can include user-defined fields (UDFs) in your query. By specifying a UDF value of true, you indicate to the API that the field you provide in your query is user-defined. The udf expression must always follow the field expression in the API call. Including the UDF value is unnecessary if you are not calling a user-defined field.

 ```json
  "filter": [
        {
            "op": "SelectedOperator",
            "field": "NameofField",
            "udf": true,
            "value": "DesiredValue"
        }
```

## Notes about QueueID

QueueID is a picklist in tickets (not a database table reference)

* 5 = Client Portal (DO NOT USE)
* 8 = Monitoring Alert
* 10 = Scheduled Tasks



[def]: #configuration-constants