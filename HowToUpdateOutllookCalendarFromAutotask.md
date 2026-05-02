# updating Calendars with Autotask info

 there are two styles, either via 365 (needs MgGraph module) or Outlook (needs outlook running)

  before the Update-CalendarFGRomTickets command can be run

- the autotask API login must have been run once (but this does not need to be run every session - just ONCE ever)
- the Set-AutoTaskEngineerIDToReport  command needs to be run and the email account (or your actual ID if you know it) assigned to YOUR autotask account assigned.  this will store a record of you autotaskID - so that the below functions can use it

```powershell
  Set-AutoTaskEngineerIDToReport -email sean.macey@imatec.co.nz 
```

## 365 MgGraph aaproach

If using MgGraph 365 connection. First install MgGraph

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force

#then run the the update (and if needed accept the scope)
Update-365CalendarFromTickets -LastXWeeks 1 
Update-365CalendarFromTickets -LastXWeeks 1 -UpdateExistingInDetail
Update-365CalendarFromTickets -LastXWeeks 1 -DoNotProvideDetaledInfoOnNew
```

## Outlook sync Approach

In order for outlook function to work you must also have OUTLOOK OPEN
  
```powershell
Update-CalendarFromTickets  -LastXWeeks 1 -ProvideDetaledInfo
```

to update/create outlook calendar entries fro all autotask ticket entries in the last week. the -ProvideDetailedInfo param will cause the update to take much longer but will also ensure ticket number, Company Name and ticket Title are included in the entry

```powershell
Update-CalendarFromTickets  -LastXWeeks 1 -UpdateExistingDetailedInfo
```

every time the Update commmand is run the existing calendar entries will have their start and stop hours updated.
But if you also want to update the TicketNumber, CompanyName and Title then use -UpdateExistingDetailedInfo
