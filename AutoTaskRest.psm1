$global:kissATAPIpath = "$home\kiss-atapi"
$global:kissATAPIfile = 'kissAtapilogin.json'
[Int32]$global:ATResourceID = $null
[int[]]$script:ATnonBillableCodes = @(29682861) #Non Billable Support
[int]$script:CONSTInternalClasificationCode = 200
[int]$script:ATtaskCodesRMM = 29712660
[int[]]$script:CONSTATLeaveCodes = @(91206, 29718729)
[int[]]$script:CONSTATSickCodes = @(91207)





#region ── Module Installation ────────────────────────────────────────────────
# To install as a PowerShell 7 module, download the script from:
#   https://gitlab.kissit.co.nz/kiss/autotaskrest/-/raw/main/AutoTaskRest.ps1?inline=false
# Then run the Install-AutoTaskRestModule.ps1 script included in the repository,
# or follow the manual instructions in README.md.
#
# Module path (PowerShell 7):
#   $HOME\Documents\PowerShell\Modules\AutoTaskRest\AutoTaskRest.psm1
#
# Legacy PowerShell 5 path (not recommended):
#   C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\AutoTaskRest\AutoTaskRest.psm1
#
# After installation, set execution policy if needed:
#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
#endregion

# check for to test REST API  https://webservices6.autotask.net/ATServicesRest/swagger/ui/index
# Check for REST API information - such as entitis and calling methods and syntax
#

#generic filter to use {"filter":[{"op":"gte","field":"id","value":"0"}]}


#GET vs READ for extract
#Measure vs Build
#invoke

<#
Organization type
The organization type describes your company's relationship with another organization. Organization types are pre-defined; they cannot be modified or added to. Options include the following:

1Customer: An organization to which you are selling products or services.
2Lead: An organization type used to indicate a potential customer.
3Prospect: An organization type used to indicate a likely customer.
4Dead: A lead that never became a customer.
6Cancelation: An Autotask organization type denoting a former customer.
7Vendor: An organization type whose primary business relationship with your company is to provide goods and services.
8Partner: An organization type assigned to organizations like VARs, outsourcing partners, etc.
#>

function Convert-ObjArrayDateTimesToSearchableStrings {
    <#
    .SYNOPSIS
    Adds sortable string versions of any DateTime properties found on objects in an array.

    .DESCRIPTION
    Iterates over every object in the supplied array and inspects each property.
    For any property whose value is a [DateTime], a new property named '<OriginalName>_Searchable'
    is added containing the date formatted as 'yyyy-MM-ddTHH:mm:ss'.

    This makes date fields easily sortable and filterable as strings in CSV exports,
    PowerBI, or Where-Object comparisons — avoiding the locale-dependent formatting
    that can cause issues when working with DateTime objects directly.

    .PARAMETER items
    An array of objects to process. Modified in place.

    .EXAMPLE
    $entries = Get-ATTimeEntries -LastxMonths 1
    Convert-ObjArrayDateTimesToSearchableStrings -items $entries
    # $entries now has e.g. dateWorked_Searchable alongside dateWorked

    .NOTES
    This is a secondary implementation used in later parts of the script.
    The pipeline-enabled version earlier in the script (with -obj parameter) is the primary one.
    Objects are modified in place; nothing is returned.
    #>
    param (
        [object[]]$items
    )
    foreach ($item in $items) {
        foreach ($property in $item.PSObject.Properties) {
            if ($property.Value -is [DateTime]) {
                $item | Add-Member -NotePropertyName ($property.Name + "_Searchable") -NotePropertyValue ($property.Value.ToString("yyyy-MM-ddTHH:mm:ss")) -Force
            }
        }
    }
}

# function ConvertTo-EscapedString() {
#     <#
#     .SYNOPSIS
#     Escapes special characters in a string for safe use in Autotask API URL query strings.

#     .DESCRIPTION
#     Replaces characters that would break a URL query string (such as '&') with their
#     percent-encoded equivalents. Currently encodes '&' as '%26'.
#     Use this on any user-supplied string before embedding it in an Autotask API filter.

#     .PARAMETER inputString
#     The raw string to be escaped. Mandatory.

#     .EXAMPLE
#     ConvertTo-EscapedString -inputString "Smith & Sons"
#     Returns: "Smith %26 Sons"

#     .NOTES
#     Only '&' is currently escaped. Extend the function body if other special characters
#     (e.g. '#', '+') are encountered in company names or other fields.
#     #>
#     param (
#         [Parameter(Mandatory = $true)]
#         [string]
#         $inputString
        
#     )
#     if ($inputString -like "*&*") { $inputString = $inputString -replace '&', '%26' }
#     #$inputString = [System.Uri]::UnescapeDataString($inputString)
#     return $inputString
# }

function Export-ATCompanies() {
    <#
    .SYNOPSIS
    Exports Autotask company data (and classification icons) to CSV files.

    .DESCRIPTION
    Writes two CSV files to the specified path (or the current directory):
      KissAtClassificationIcons.csv  — all classification icon IDs and names
      KissAtCompanies.csv            — all companies (including inactive) with
                                       primary/secondary engineer assignments

    This function calls Get-ATCompanies with -includeInactive and -GetEngineers,
    which performs multiple API round-trips and can take approximately 3 minutes.

    .PARAMETER exportType
    Output format. Currently only CSV is supported. Default is CSV.

    .PARAMETER path
    Optional output directory. A trailing backslash is appended automatically.
    If omitted, files are written to the current working directory.

    .EXAMPLE
    Export-ATCompanies

    .EXAMPLE
    Export-ATCompanies -path "C:\Reports"

    .NOTES
    JSON export is defined in the switch but not yet implemented.
    Expect a run time of around 3 minutes due to the engineer-assignment lookups.
    #>
    [CmdletBinding()]
    param (
        # Parameter help description
        #[Parameter(AttributeValues)]
        [ValidateSet("CSV", "JSON")]
        [string]
        $exportType = "CSV",
        [string]
        $path 
    )
    if ($path) { $path = "$path\\" }
    write-host "Export-ATCompanies will take about 3 minutes to run!"
    switch ($exportType) {
        "CSV" {
            write-host "Export-ATCompanies =>Exporting ClassificationIcons"
            Invoke-ATQuery -entityName 'v1.0/ClassificationIcons' -includeFields "id", "name" | export-csv "$($path)KissAtClassificationIcons.csv" -NoTypeInformation -Force
            write-host "Export-ATCompanies =>Exporting Companies"
            Get-ATCompanies -includeInactive -GetEngineers | export-csv "$($path)KissAtCompanies.csv" -NoTypeInformation -Force
        }
        default {

        }

    }
    write-host "Done Export-ATCompanies" -ForegroundColor green
}



function Export-ATTickets() {
    <#
    .SYNOPSIS
    Exports Autotask ticket data to CSV files.

    .DESCRIPTION
    Exports two sets of tickets to CSV:
      1. All currently open (non-completed) tickets with an assigned resource.
      2. Optionally, tickets that have had activity within the last N months.

    Output files are written to the current directory (or to the path specified by -path):
      - TicketsNotCompleted.csv  — all open tickets
      - TicketsActioned.csv      — tickets active within WhereLastActionOccurWithinLastMonths (if > 0)

    .PARAMETER WhereLastActionOccurWithinLastMonths
    How many months back to include recently-actioned tickets in TicketsActioned.csv.
    Default is 0 (disabled — only the non-complete export runs).

    .PARAMETER path
    Optional output directory path. A trailing backslash is appended automatically.
    If omitted, files are written to the current working directory.

    .EXAMPLE
    Export-ATTickets
    # Exports only TicketsNotCompleted.csv to the current directory

    .EXAMPLE
    Export-ATTickets -WhereLastActionOccurWithinLastMonths 3 -path "C:\Reports"
    # Exports both TicketsNotCompleted.csv and TicketsActioned.csv to C:\Reports\

    .NOTES
    Calls Get-ATTickets internally. Can be slow for large ticket sets.
    #>
    param (
        [Parameter(Mandatory = $false)]
        [int]
        $WhereLastActionOccurWithinLastMonths = 0,
        [string]$path
    )

    if ($path) { $path = "$path\\" }
    New-Item -ItemType Directory -Name data -ErrorAction SilentlyContinue | Out-Null
    if ($WhereLastActionOccurWithinLastMonths -gt 0) {
        Get-ATTickets -LastxWeeks $null -LastActionFromDate (Get-Date).AddMonths(-$WhereLastActionOccurWithinLastMonths) | Export-csv "$($path)TicketsActioned.csv" -NoTypeInformation -Force
    }
    Get-ATTickets -IncludeAllNonComplete | Export-csv "$($path)TicketsNotCompleted.csv" -NoTypeInformation -Force

}


function Export-ATTimeRecords() {
    <#
    .SYNOPSIS
    Exports time records and related reference data to multiple CSV or JSON files.

    .DESCRIPTION
    Retrieves time entries for the last N months (or as specified) together with
    the tickets active in the same window, billing codes, and engineer data, then
    writes them as a set of files suitable for Power BI or other reporting tools.

    Files written (CSV mode):
      KissAtClassificationIcons.csv  — not written here; see Export-ATCompanies
      KissBillingCodes.csv
      KissEngineers.csv
      KissTimeEntries.csv
      KissDaily.csv                  — per-engineer daily stats from Get-ATDailyTimeStats
      KissTickets.csv

    .PARAMETER LastxMonths
    How many months back to retrieve time entries from. Default is 3.

    .PARAMETER exportType
    Output format. Currently only CSV is fully supported. Default is CSV.

    .PARAMETER path
    Optional output directory. A trailing backslash is appended automatically.
    If omitted, files are written to the current working directory.

    .EXAMPLE
    Export-ATTimeRecords

    .EXAMPLE
    Export-ATTimeRecords -LastxMonths 6 -path "W:\autotask"

    .NOTES
    JSON export is present in the default branch but has not been fully validated.
    This function can take several minutes on large datasets.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]
        $LastxMonths = 3,
        # Parameter help description
        #[Parameter(AttributeValues)]
        [ValidateSet("CSV", "JSON")]
        [string]
        $exportType = "CSV",
        [string]$path
    )
    Write-Host "Export-ATTimeRecords: will take some time to run"
    write-host " Export-ATTimeRecords =>preparing Time Entries"

    $i = Get-ATTimeEntries -LastxMonths $LastxMonths  -IncludeSummaryNotes #-includeTicketDetails -includeEngineerDetails -IncludeBillingDetails

    $earliestDate = ($i | Measure-Object dateWorked -min).Minimum
    write-host " Export-ATTimeRecords =>preparing Ticket Details"

    $Tickets = Get-ATTickets -LastActionFromDate $earliestDate
    if ($path) { $path = "$path\\" }

    switch ($exportType) {
        "CSV" {
            write-host " Export-ATTimeRecords : =>Billing Codes"

            Invoke-ATQuery -entityName 'v1.0/BillingCodes' | Export-csv "$($path)KissBillingCodes.csv" -NoTypeInformation -Force
            write-host " Export-ATTimeRecords : =>Resources (Engineers) and timeEntries"
            Get-ATEngineers | export-csv "$($path)KissEngineers.csv" -NoTypeInformation -Force
            
            $i | export-csv "$($path)KissTimeEntries.csv" -NoTypeInformation -Force
            write-host " Export-ATTimeRecords =>DailyTime Stats and tickets"
            Get-ATDailyTimeStats -TimeEntries $i | Export-Csv "$($path)KissDaily.csv" -NoTypeInformation -Force
            $Tickets | Export-Csv "$($path)KissTickets.csv" -NoTypeInformation -Force
 
            #  Invoke-ATQuery -entityName 'v1.0/ResourceTimeOffBalances' | Export-csv ResourceTimeOffBalances.csv -NoTypeInformation

            #Holiday and Holidayset records not in use
        }
        default {
            Invoke-ATQuery -entityName 'v1.0/BillingCodes' | ConvertTo-Json | Out-File -FilePath KissBillingCodes.json -Force
            Get-ATEngineers | ConvertTo-Json | Out-File -FilePath KissEngineers.json -Force
            # $i = Get-ATTimeEntries -LastxMonths $LastxMonths
            $i | ConvertTo-Json | Out-File -FilePath  KissTimeEntries.json -Force
            Get-ATDailyTimeStats -TimeEntries $i | ConvertTo-Json | Out-File -FilePath  KissEnginerDailies.json -Force
            $Tickets | ConvertTo-Json | Out-File -FilePath  KissTickets.json -Force
            #    Invoke-ATQuery -entityName 'v1.0/ResourceTimeOffBalances' | Out-File -FilePath ResourceTimeOffBalances.json

        }

    }
    write-host "Done Export-ATTimeRecords" -ForegroundColor green
}

function Find-ATCompaniesInTickets() {
    <#
    .SYNOPSIS
    Returns the company records for all companies referenced in a ticket collection.

    .DESCRIPTION
    Groups the supplied ticket array by CompanyID, then calls Get-ATCompanies once
    per unique CompanyID and yields each company object.

    Useful when you have a set of tickets and want to enrich them with full company
    detail without fetching all companies.

    Note: this can be slow for large ticket sets (e.g. ~3 minutes per 100 companies)
    because it makes one API call per unique company.

    .PARAMETER tickets
    An array of ticket objects, as returned by Get-ATTickets.
    Each ticket must have a CompanyID property.

    .EXAMPLE
    $tickets = Get-ATTickets -IncludeAllNonComplete
    $companies = Find-ATCompaniesInTickets -tickets $tickets

    .NOTES
    For large sets consider using Get-ATCompanies -id $uniqueIDs instead, which
    batches the lookup into chunks of 100 and is significantly faster.
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [object[]]
        $tickets
    )
    $companies = $tickets | Group-Object CompanyID
    foreach ($companyID in $companies) {
        $company = (Get-ATCompanies -id $companyID.Name | Select-Object -First 1)
        $company
    }

}


function Get-AT365CalendarEvents {
    <#
    .SYNOPSIS
    Retrieves Microsoft 365 calendar events via Microsoft Graph for a specified date range.

    .DESCRIPTION
    Connects to Microsoft Graph with Calendars.ReadWrite scope and returns calendar events
    starting from the Sunday of LastXWeeks weeks ago (or from FromDateLocal if supplied).
    Events are returned with their singleValueExtendedProperties expanded so the Autotask
    TimeEntry ID custom property can be read.

    When -OnlyGetItemsWithAutotaskTimeEntryIDsInBody is used, only events whose body
    contains a 'TimeEntry: <number>' pattern are returned, and each is enriched with an
    AutotaskTimeEntryID property for easy matching.

    .PARAMETER LastXWeeks
    How many weeks back to retrieve events from. Default is 1.
    Ignored if FromDateLocal is supplied.

    .PARAMETER OnlyGetItemWithAutotaskTimeEntryIDsInBody
    Switch. When set, filters results to only events that contain an Autotask TimeEntry ID
    in their body content, and adds an AutotaskTimeEntryID property to each.

    .PARAMETER FromDateLocal
    Optional. A specific local DateTime to use as the start of the retrieval window,
    overriding the LastXWeeks calculation.

    .EXAMPLE
    Get-AT365CalendarEvents -LastXWeeks 4

    .EXAMPLE
    Get-AT365CalendarEvents -LastXWeeks 2 -OnlyGetItemsWithAutotaskTimeEntryIDsInBody

    .NOTES
    Requires the Microsoft.Graph.Calendar module and an active Connect-MgGraph session
    with Calendars.ReadWrite scope.
    #>
    [CmdletBinding()]
    param (
        [int]$LastXWeeks = 0,
        [switch]$OnlyGetItemWithAutotaskTimeEntryIDsInBody = $false,
        [System.Nullable[DateTime]]$FromDateLocal = $null
    )

    if ($null -eq $FromDateLocal) {
        $FromDateLocal = (Get-ATWeekStart -LastXWeeks $LastXWeeks)
    }
    # $CURRENTDATE = GET-DATE -Hour 0 -Minute 0 -Second 0
    $DateTostartCalCheck = $FromDateLocal.ToUniversalTime().ToString("o") # ISO 8601 format for date-time
    #$DateTostartCalCheck = $CURRENTDATE.AddDays(-7 * ($LastXWeeks + 1)).ToUniversalTime().ToString("o") # ISO 8601 format for date-time
    write-verbose "Get-AT365CalendarEvents: will check calendar items starting from UTC $DateTostartCalCheck" #-ForegroundColor Green
    if (-not(Get-Module -ListAvailable -Name  Microsoft.Graph)) { 
        # if (-not(Get-InstalledModule Microsoft.Graph)) { 
        #Get-Module -ListAvailable -Name Microsoft.Graph  
        Write-Host "Microsoft Graph module not found" -ForegroundColor Black -BackgroundColor Yellow
        $install = Read-Host "Do you want to install the Microsoft Graph Module?"
  
        if ($install -match "[yY]") {
            Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
        }
        else {
            Write-Host "Microsoft Graph module is required." -ForegroundColor Black -BackgroundColor Yellow
            throw "Microsoft Graph module is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
        } 
    }
    Connect-MgGraph -Scopes  "Calendars.ReadWrite"
    #Select-MgProfile -Name beta  # optional but recommended for calendar precision
   
    $365me = (Get-MgContext).Account
    write-verbose "Get-AT365CalendarEvents: connected to Microsoft Graph as $365me, now getting calendar events from the last $LastXWeeks weeks"

    $events = Get-MgUserEvent -UserId $365me -All -Filter "start/dateTime ge '$DateTostartCalCheck'" -ExpandProperty "singleValueExtendedProperties(`$filter=id eq 'String {1e388ea9-5c0d-4aec-aaf9-8150a6e7797c} Name AutotaskTimeEntryID'`)"
    
    write-verbose "Get-AT365CalendarEvents: retrieved $($events.Count) calendar events from Microsoft Graph, now looking for events with Autotask Time Entry IDs in the body"
    if ($OnlyGetItemWithAutotaskTimeEntryIDsInBody) {
        # $events = $events | Where-Object { $_.Body.Content -match 'TimeEntry:\s*(\d+)' }
        $Calevents = @()   
        foreach ($anevent in $events) {
            if ($anevent.Body.Content -match 'TimeEntry:\s*(\d+)') {
                $number = $Matches[1]
                if ($number -gt 0) {
                    write-verbose "Get-AT365CalendarEvents: Found event with TimeEntry: $([int]$number) in the body. Subject: $($anevent.Subject) Start: $($anevent.Start.DateTime) End: $($anevent.End.DateTime)" #-ForegroundColor Cyan
                    $anevent | Add-Member -NotePropertyName AutotaskTimeEntryID -NotePropertyValue ([int]$number) -Force
                    $Calevents += $anevent                
                }

            }
        }
        write-verbose "Get-AT365CalendarEvents: filtered down to $($Calevents.Count) events that have Autotask Time Entry IDs in the body"
        return $Calevents
    }
    $events
    # foreach ($event in $events) {
    #     write-host "Get-AT365CalendarEvents: Event: $($event.Subject) Start: $($event.Start.DateTime) End: $($event.End.DateTime) Body: $($event.Body.Content)"
    #     Update-MgUserEvent -UserId $365me -EventId $event.Id  -BodyParameter @{Body = $event.Body.Content + "`n`nThis event was retrieved and updated by the Get-AT365CalendarEvents function" }
    #     return
    # }   


}

function Get-ATBillingCodes() {
    <#
    .SYNOPSIS
    Returns all billing code records from Autotask.

    .DESCRIPTION
    Retrieves every billing code defined in Autotask (v1.0/BillingCodes).
    Billing codes are used to categorise time entries (e.g. Normal Support, After Hours,
    Leave, Sick, RMM, Training). The returned objects include the id and name fields
    needed to interpret billingCodeID values on time entries.

    .EXAMPLE
    $codes = Get-ATBillingCodes
    $codes | Select-Object id, name | Sort-Object id

    .NOTES
    No parameters. Returns all billing codes without pagination (typically a small dataset).
    #>
    [CmdletBinding()]
    param (
    )
    Write-verbose "Polling Autotask for Billing Codes" #-ForegroundColor Green"
    Invoke-ATQuery -entityName 'v1.0/BillingCodes' 

}

function Get-ATClassificationIcons() {
    <#
    .SYNOPSIS
    Returns all company classification icon records from Autotask.

    .DESCRIPTION
    Retrieves the full list of ClassificationIcons from the Autotask API.
    Each record contains an id and a description/name that maps a numeric
    classification code (e.g. 7 = Vendor, 200 = Internal) to a human-readable label.
    This is typically called internally by Get-ATCompanies to enrich company
    records with a ClassificationDetails field.

    .EXAMPLE
    $icons = Get-ATClassificationIcons
    $icons | Select-Object id, name

    .NOTES
    No parameters. Returns all icons; they are relatively few and are not paginated.
    #>
    [CmdletBinding()]
    param (  )
    $rc = Invoke-ATQuery -entityName 'v1.0/ClassificationIcons'   -SearchFirstBy id
    $rc
}

function Get-ATCompanies {
    <#
    .SYNOPSIS
    returns a list of companies (or just one of)
    takes a long while to run if there are many customers
    
    .DESCRIPTION
     returns a list of companies (or just one of)

    
    .PARAMETER id
    company ID specific serach
    
    .PARAMETER CompanyName
    search for a name (by default any close matches are returned
    
    .PARAMETER includeFields
    Parameter description
    
    .PARAMETER exactNameMatch
    if used then only the exact match for the company name is returned
    
    .PARAMETER includeInactive
    ensures that even inactive clients are returned
    default is NO
    
    .PARAMETER GetEngineers
    also add information about the Primary and Secondary engineers
    
    .EXAMPLE
    Get-ATCompanies

     Get-ATCompanies -CompanyName "imatec" -debug 
        DEBUG: getiing  Companies items  https://webservices6.autotask.net/atservicesrest/v1.0/Companies/query?search={"IncludeFields":["id", "isActive","companyName","companyType","classification","lastActivityDate", "Branch"],"filter":[{"op":"contains","Field":"companyName","value":"imatec"}]}  

        id               : 29762985
        classification   : 7
        companyName      : Imatec Solutions (As Customer)
        companyType      : 1
        isActive         : True
        lastActivityDate : 2023-08-01T05:27:43
        Branch           : Matamata

        id               : 29762986
        classification   : 1
        companyName      : Imatec - Test Customer
        companyType      : 1
        isActive         : True
        lastActivityDate : 2022-04-23T07:39:24
        Branch           : Matamata


    
    .NOTES
    General notes
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByID')]
    param (
        [Parameter(ParameterSetName = 'ByID', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [int[]]
        $id = @(),
        # Parameter help description
        [Parameter(ParameterSetName = 'ByName', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $CompanyName,

        # Parameter help description
        #[Parameter(AttributeValues)]
        [string]
        $includeFields = '"id", "isActive","companyName","companyType","classification","lastActivityDate", "Branch"' ,

        # Parameter help description
        #[Parameter(AttributeValues)]
        [switch]
        $exactNameMatch,

        # Parameter help description
        #[Parameter(AttributeValues)]
        [switch]
        $includeInactive = $false,

        # Parameter help description
        #[Parameter(AttributeValues)]
        [switch]
        $GetEngineers = $false,
        # Parameter help description
        [Parameter(Mandatory = $false)]
        [switch]
        $DontExpandChildIDFields = $false

    )

 
    if ($exactNameMatch) { $op = "eq" } else { $op = "contains" } 
    
    switch ($true) {
        { $id.count -gt 0 } {
            # write-verbose "Get-ATCompanies - for a ID $id"

            if ( $id ) {
                if ($id.count -eq 1) {
                    $ida = $id[0]
                    write-verbose "Get-ATCompanies - for a single ID $ida"
                    $rc = Invoke-ATQuery -entityName 'v1.0/Companies' -id $ida -SearchFirstBy id -includeFields $includeFields

                    #  $result += Invoke-ATQuery -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFirstBy id -ID $id  -isActive:$isActive
                }
                else {
                    write-verbose "Get-ATCompanies - for a ID SET $id "
                    $chunkSize = 100
                    $rc = @()
                    for ($i = 0; $i -lt $id.Count; $i += $chunkSize) {
                        $chunk = [int[]]$id[$i..([Math]::Min($i + $chunkSize - 1, $id.Count - 1))]
                        [string]$cci = $chunk -join ','
                        $searchFilter1 = '{"op":"in","Field":"id","value":[' + $cci + ']}'
                        if ($searchFilter1) {
                            $rc += Invoke-ATQuery -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy  Nothing  -SearchFurtherBy $searchFilter1
                        }
                    }
                }
                #  $result += Invoke-ATQuery -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFirstBy  Nothing '{"op":"in","Field":"id","value":[' + ($id -join ", ") + "]}"  

            }   
            break
        }
        { $CompanyName } {
            write-verbose "Get-ATCompanies - for a exact match :$companyName"
            $escapedCompanyName = [Uri]::escapeDataString( $companyName )
 #          $escapedCompanyName = convertto-escapedString -inputString $companyName
            [string]$srch = "{""op"":""$op"",""Field"":""companyName"",""value"":""$escapedCompanyName""}"  #{"op":"contains","Field":"companyName","value":"imatec"}
            $rc = Invoke-ATQuery -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy Nothing  -SearchFurtherBy $srch
            break 
        }
        { $includeInactive -eq $true } { 
            write-verbose "Get-ATCompanies - for ALL companies including inactive"
            $rc = Invoke-ATQuery -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy id -CheckDuplicatesOf "id"
            break 
        }
        default {
            write-verbose "Get-ATCompanies - for ALL Active companies"
            $rc = Invoke-ATQuery -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy isActive -CheckDuplicatesOf "id"
        }
    }

    if ($rc) {
        $branchi = ($rc.userDefinedFields | Where-Object { $_.name -eq "Branch" })
        $branch = $null
        if ($branchi) { $branch = $branchi[0] }
        $rc = $rc | select-Object -Property * , @{name = "Branch"; e = { $branch.value } } -ErrorAction SilentlyContinue | Select-Object -ExcludeProperty userDefinedFields
        if (!($DontExpandChildIDFields -eq $true)) {

            
            Convert-ObjArrayDateTimesToSearchableStrings -obj $rc #|Out-Null

            #$rc.userDefinedFields
            #$rc = $rc | Select-Object -ExcludeProperty userDefinedFields

            if ($GetEngineers) {
                $rc | Add-Member -NotePropertyName Primary -NotePropertyValue ""
                $rc | Add-Member -NotePropertyName Secondary -NotePropertyValue ""
                $AllPrimeTechnicians = Get-ATCompanyEngineers
                # this updates the objects in $array1
                foreach ($i in $rc) {
                    $thisprime = $AllPrimeTechnicians | Where-Object CompanyID -eq $i.id | Select-Object -First 1
                    if ($thisprime) {
                        $i.primary = $thisprime.primary
                        $i.secondary = $thisprime.secondary
                    }
                }
            }
            
            # get special comments about company including whether Residential or commercial
            $classificationIcons = Get-ATClassificationIcons
            $rc | Add-Member -NotePropertyName 'ClassificationDetails' -NotePropertyValue "" -Force
            if ($classificationIcons) {
                $CompanyGroups = $rc | Where-Object classification | Group-Object classification
    
                foreach ($item in $CompanyGroups ) {
                    if ($item.name) {
                        $classificationDetail = ($classificationIcons | Where-Object id -eq ($item.name)).description
                        $item.group | Add-Member -NotePropertyName 'ClassificationDetails' -NotePropertyValue "$classificationDetail" -Force
                    }

                }
            }
        }
        write-verbose "Done Get-ATCompanies" #-foregroundColor Green

        return $rc
    }
}

function Get-ATCompanyAlert() {
    <#
    .SYNOPSIS
    Returns the alert text for a specific alert type on a given Autotask company.

    .DESCRIPTION
    Queries the Autotask CompanyAlerts endpoint for a single company and alert type,
    then returns the alertText string of the first matching record.
    Alert type 1 (the default) is typically used for primary/secondary engineer assignments.

    .PARAMETER AlertTypeID
    The numeric alert type to search for.
    Default is 1 (Primary/Secondary engineer alert).

    .PARAMETER CompanyID
    The Autotask company ID to query. Mandatory.

    .EXAMPLE
    Get-ATCompanyAlert -CompanyID 29762985
    # Returns the text of alert type 1 for company 29762985

    .EXAMPLE
    Get-ATCompanyAlert -CompanyID 29762985 -AlertTypeID 3

    .NOTES
    Returns only the alertText of the first matched record, not the full alert object.
    Use Get-ATCompanyChildAlerts to retrieve all alerts for a company.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $CompanyID,
        [Parameter(Mandatory = $false)]
        [int]
        $AlertTypeID = 1
    )
    Write-verbose "Polling Autotask for CompanyID $CompanyID and AlertTypeID $AlertTypeID  "
    $u = Invoke-ATQuery -entityName 'v1.0/CompanyAlerts' -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""alertTypeID"",""value"":""$AlertTypeID""},{""op"":""eq"",""Field"":""CompanyID"",""value"":""$CompanyID""}" # -Verbose
    if ($u) {
        $u[0].alertText
    }
}


function Get-ATCompanyChildAlerts() {
    <#
    .SYNOPSIS
    Returns all alert records associated with a specific Autotask company.

    .DESCRIPTION
    Queries the Autotask Companies/{CompanyID}/Alerts child endpoint and returns
    every alert item for that company, regardless of alert type.
    Used internally by Set-ATCompanyEngineers to retrieve existing alerts
    before deciding whether to create, update, or delete them.

    .PARAMETER CompanyID
    The numeric Autotask company ID. Mandatory.

    .EXAMPLE
    $alerts = Get-ATCompanyChildAlerts -CompanyID 29762985
    $alerts | Select-Object alertTypeID, alertText

    .NOTES
    Returns $u.items — the raw items array from the Autotask response.
    Returns $null if no alerts exist for the company.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [int]
        $CompanyID
    )
    Write-verbose "Polling Autotask for CompanyID $CompanyID for all its alerts"
    $u = Invoke-ATQuery  -UrlFixedSuffix "v1.0/Companies/$CompanyID/Alerts" 
    if ($u) {
        $u.items
    }
 
}

function Get-ATCompanyEngineers() {
    <#
    .SYNOPSIS
    Returns primary and secondary technician assignments for all Autotask customers.

    .DESCRIPTION
    Queries the Autotask CompanyAlerts endpoint for alert type 1 (the engineer-assignment
    alert) and parses the alertText of each record to extract the Primary and Secondary
    engineer names stored there by Set-ATCompanyEngineers.

    When -IncludeCompanyDetail is set, additional company fields (name, branch, isActive,
    classification, last activity) are joined from Get-ATCompanies.

    When -UnassignInactiveCustomers is set (implies -IncludeCompanyDetail), Primary and
    Secondary are blanked for any company that is no longer active.

    .PARAMETER alertTypeID
    The Autotask alert type to query. Default is 1 (primary/secondary engineer alert).

    .PARAMETER IncludeCompanyDetail
    Switch. Enriches each result with company name, branch, isActive, classification,
    and last activity date from Get-ATCompanies.

    .PARAMETER UnassignInactiveCustomers
    Switch. When set, Primary and Secondary are cleared for inactive companies.
    Implies -IncludeCompanyDetail.

    .EXAMPLE
    Get-ATCompanyEngineers

    .EXAMPLE
    Get-ATCompanyEngineers -IncludeCompanyDetail

    .EXAMPLE
    Get-ATCompanyEngineers -UnassignInactiveCustomers

    .NOTES
    Engineer names are parsed from free-text alertText using regex; formatting in the
    alert must follow the "Primary Tech: Name" / "Secondary Tech: Name" convention set
    by Set-ATCompanyEngineers for parsing to succeed.
    #>
    
    #Get prime and secondary
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory = $false)]
        [int]
        $alertTypeID = 1, #could be 1,2,3
        # Parameter help description
        [Parameter(Mandatory = $false)]
        [switch]
        $IncludeCompanyDetail = $false,
        # Parameter help description
        [Parameter(Mandatory = $false)]
        [switch]
        $UnassignInactiveCustomers = $false
    )


    Write-Host "Polling Autotask for Company(Client) Prime and (Secondary) Engineers"
    $u = Invoke-ATQuery -entityName 'v1.0/CompanyAlerts' -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""alertTypeID"",""value"":""$alertTypeID""},{""op"":""contains"",""Field"":""alertText"",""value"":""primary""}" # -Verbose
    # [System.Object[]]$PrimeTechnicians = $null
    if ($UnassignInactiveCustomers -eq $true) { $IncludeCompanyDetail = $true }
    foreach ($l in $u) {
        if ($IncludeCompanyDetail -eq $true) {
            $assignedTech = [PSCustomObject]@{
                CompanyID      = $l.CompanyID
                Company        = ""
                Primary        = $null
                Secondary      = $null
                Branch         = ""
                isActive       = $False
                LastAction     = ""
                Classification = ""
            }
            $classifications = Get-ATClassificationIcons
        }
        else {
            $assignedTech = [PSCustomObject]@{
                CompanyID = $l.CompanyID
                Primary   = $null
                Secondary = $null
                # TextPrimary    = ""
                # TextSecondary  = ""
                # CompanyAlertID = $null
            }
        }

        if ($l.AlertText -imatch "secondary\s+tech.*[:][\s|\w]*\n|secondary\s+engineer.*[:][\s|\w]*\n|secondary\s+tech.*[:][\s|\w]*|secondary\s+engineer.*[:][\s|\w]*") {
            $assignedTech.secondary = ($Matches[0]) -replace ("\n", "")
            #$assignedTech.CompanyAlertID = $l.ID
            $assignedTech.secondary = $assignedTech.secondary -ireplace [regex]::Escape("secondary"), ""
            $assignedTech.secondary = $assignedTech.secondary -ireplace [regex]::Escape("engineer"), ""
            $assignedTech.secondary = $assignedTech.secondary -ireplace [regex]::Escape("tech"), ""
            $assignedTech.secondary = $assignedTech.secondary.replace(":", "").trim()
        } 

        if ($l.AlertText -imatch "primary\s+tech.*[:][\s|\w]*\n|primary\s+engineer.*[:][\s|\w]*\n|primary\s+tech.*[:][\s|\w]*|primary\s+engineer.*[:][\s|\w]*") {
            $assignedTech.Primary = ($Matches[0]) -replace ("\n", "") 
            #$assignedTech.CompanyAlertID = $l.ID
            $assignedTech.Primary = $assignedTech.Primary -ireplace [regex]::Escape("primary"), ""
            $assignedTech.Primary = $assignedTech.Primary -ireplace [regex]::Escape("engineer"), ""
            $assignedTech.Primary = $assignedTech.Primary -ireplace [regex]::Escape("tech"), ""
            $assignedTech.Primary = $assignedTech.Primary.replace(":", "").trim()

        }




       

        if ($assignedTech.Primary -or $assignedTech.Secondary) {
            # we found a RECORD for primary/secondary in AutoTask
            if ($IncludeCompanyDetail) {
               
                $company = Get-ATCompanies -id $assignedTech.CompanyID -DontExpandChildIDFields
                $assignedTech.Company = $company.companyName
                if ($company.classification) {
                    $assignedTech.Classification = ($classifications | where-object id -eq $company.classification).name
                }
                $assignedTech.Branch = $company.Branch
                $assignedTech.isActive = $company.isActive
                if (($UnassignInactiveCustomers -eq $true) -and !($assignedTech.isActive -eq $true)) {
                    $assignedTech.Primary = ""
                    $assignedTech.Secondary = ""
                }
                $assignedTech.LastAction = ($company.lastActivityDate -split (" "))[0]
            }
            #$PrimeTechnicians += $assignedTech
            $assignedTech
        }
    }
    Write-Host "DONE Polling Autotask for Company(Client) Prime and (Secondary) Engineers"
    # return $PrimeTechnicians
}

function Get-ATContacts() {
    <#
    .SYNOPSIS
    get an array of contacts from Autotask
    
    .DESCRIPTION
    get an array of contacts from Autotask, by ID , or emailaddress
    
    .PARAMETER ID
    the autotask ID of a contact
    
    .PARAMETER eMail
    the emailaddress of a contact(s)
    
    .EXAMPLE
    $a = Get-ATContacts -id 29692052
    
    .NOTES
    General notes
    #>
   
    [CmdletBinding(DefaultParameterSetName = 'ByID')]
    param (
        [Parameter(ParameterSetName = 'ByID', Position = 0, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [int]
        $ID = -1,
        [Parameter(ParameterSetName = 'ByEmail', Position = 0, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $eMail
    )
    begin {
        
    }
    process {
        switch ($true) {
            { $ID -ge 0 } { 
                Write-Host "Polling Autotask for Contact with ID of $ID"
                Invoke-ATQuery -entityName 'v1.0/Contacts' -SearchFirstBy id -ID $ID
                break
            }
            { $eMail.Length -gt 0 } {
                Write-Host "Polling Autotask for Contact with email of $eMail"
                Invoke-ATQuery -entityName 'v1.0/Contacts' -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""emailAddress"",""value"":""$eMail""}"
                break
            }
            Default {
                Write-Host "Polling Autotask for all Contacts"
                Invoke-ATQuery -entityName 'v1.0/Contacts' 
            }
            #     }
            # if ($ID -ge 0 ) {
            #     return
            # }
            # if ($eMail) {
            #     Invoke-ATQuery -entityName 'v1.0/Contacts' -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""emailAddress"",""value"":""$eMail""}"
            #     return
            # }

        }
    }
    end {}
}

function Get-ATCredentialHeader {
    <#
    .SYNOPSIS
    Private helper. Builds the Autotask REST API authentication header from saved or supplied credentials.

    .DESCRIPTION
    Loads credentials from the kiss-atapi login file (or from a supplied LoginInfo object),
    decrypts the secret and API integration code using DPAPI, builds the header hashtable,
    then immediately zeros the BSTR memory buffers so the plain-text secret exists in memory
    for the shortest possible time.

    Returns a hashtable ready to pass as -Headers to Invoke-RestMethod, and the resolved
    base URL as a second output value via a [ref] parameter.

    All credential decryption in the module flows through this single function so that
    security fixes only need to be made in one place.

    .PARAMETER LoginInfo
    Optional PSCustomObject with properties: url, UserName, Secret (DPAPI-encrypted string),
    atapi (DPAPI-encrypted string). When omitted the saved login file is used.

    .PARAMETER BaseUrl
    [ref] string. On return contains the resolved Autotask base URL.

    .NOTES
    Private function — not intended to be called directly by end users.
    #>
    [CmdletBinding()]
    param(
        [PSCustomObject]$LoginInfo,
        [ref]$BaseUrl
    )

    if ($LoginInfo) {
        $saveobj = $LoginInfo
    }
    elseif (Test-Path -Path "$kissATAPIpath\$kissATAPIfile") {
        $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
        if ($jsn) { $saveobj = $jsn | ConvertFrom-Json }
    }

    if (-not $saveobj -or -not $saveobj.url -or -not $saveobj.secret -or -not $saveobj.username -or -not $saveobj.atapi) {
        throw "Get-ATCredentialHeader: Credentials are missing or incomplete. Run Set-ATLogin first."
    }

    # Decrypt Secret (password)
    $securePwd = $saveobj.Secret | ConvertTo-SecureString
    $bstrPwd = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd)
    $plainPwd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrPwd)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstrPwd)

    # Decrypt atapi integration code.
    # Backward-compatibility: older login files stored atapi as plain text (no ConvertFrom-SecureString).
    # A DPAPI-encrypted string is always longer than 50 chars and contains only hex characters.
    # A plain-text API ID is short alphanumeric. We detect which format it is and handle both.
    $plainApi = $null
    try {
        $secureApi = $saveobj.atapi | ConvertTo-SecureString   # succeeds if it was DPAPI-encrypted
        $bstrApi = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureApi)
        $plainApi = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrApi)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstrApi)
    }
    catch {
        # atapi is stored as plain text (old format) — use it directly and warn once
        Write-Warning "Get-ATCredentialHeader: The API integration code in the login file is stored as plain text. Run Set-ATLogin to re-save it in encrypted form."
        $plainApi = $saveobj.atapi
    }

    if ($BaseUrl) { $BaseUrl.Value = $saveobj.url }

    $header = @{
        'ApiIntegrationCode' = $plainApi
        'UserName'           = $saveobj.UserName
        'Secret'             = $plainPwd
        'Content-Type'       = 'application/json'
    }

    # Zero out local plain-text variables immediately after building the header
    $plainPwd = $null
    $plainApi = $null

    return $header
}

function Get-ATDailyTimeStats {
    <#
    .SYNOPSIS
    calculate daily summary for each technician that is time sheeting
    requires the timeEntries object array to t=be parsed to it
    - this does not use inline processing, the timeentries must be passed as a paramneter object array
    
    .DESCRIPTION
    Long description
    creates daily expected hours, which is the greater of (normal ours worked less Leave and Sick) Or each Tech's ecpected daily hours
    
    .PARAMETER TimeEntries
    AN array of tiome entries (generated by Get-ATTimeEntries )
    
    .EXAMPLE
    Get-ATDailyTimeStats -TimeEntries $timeEntries
    
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]   
        [PSCustomObject]        $TimeEntries,
        [datetime]$UntilDate = (get-date) # check u timesheeted days for resources from earliest in toimesheet until this time - so ignore leave requests and future bookings when filling gaps
    )

    
    $culture = [System.Globalization.CultureInfo]::CreateSpecificCulture("en-NZ")
    $format = "d/MM/yyyy h:mm:ss tt"


    #$allresources = Get-ATEngineers
    # $Resources = $allresources | Where-Object { ($_.id -in $TimeEntries.resourceID) }  ## gets resources in time entries
    $timeentryEngineerIDs = $TimeEntries | select-object resourceID -Unique | select-object -ExpandProperty resourceID
    $Resources = Get-ATEngineers -id $timeentryEngineerIDs
   


    #$Resources += $allresources | Where-Object { ($_.isActive) -and ($_.DailyAvailabilities.MondayAvailableHours -or $_.DailyAvailabilities.TuesdayAvailableHours -or $_.DailyAvailabilities.WednesdayAvailableHours -or $_.DailyAvailabilities.ThursdayAvailableHours -or $_.DailyAvailabilities.FridayAvailableHours -or $_.DailyAvailabilities.SaturdayAvailableHours -or $_.DailyAvailabilities.SundayAvailableHours  ) }
    #$resourcesThatShouldTimeSheet = $Resources | Select-Object * -Unique
    
    write-verbose "Get-ATDailyTimeStats: Resources that are expected to be timesheeting $($resources.username -join (', '))"
    $LastDate = $UntilDate
    $LastDateOA = $LastDate.ToOADate()
    $StartDate = [datetime](($timeEntries | Measure-Object dateWorked -min).Minimum)
    
  




    #prepare an object array of every date between the start until the expect enddate
    $iDate = $StartDate
    [psobject[]]$datesToCheck = $null
    do {
        $oneDate = [PSCustomObject]@{
            date         = $iDate
            datestr      = $iDate.ToString('s')
            weekday      = $idate.DayOfWeek
            weekdayvalue = $idate.DayOfWeek.value__
        }
        $iDate = $iDate.AddDays(1)
        $datesToCheck += $oneDate
    }
    until ($LastDate -lt $iDate)





    #group timeentries by Resource, but ignore dates beyond the sample period (those will be leave bookings...)
    $gps = $TimeEntries | where-object  oadate -lt $LastDateOA | Group-Object resourceID #, dateWorked
    foreach ($gp in $gps) {
        #$OADate =
        [psobject[]]$OneResourceDates = $null
        #Find all resources which have time entries
        $Resource = $Resources | Where-Object { ($_.id -eq $gp.name) } | Select-Object -First 1
        

        $techDays = $gp.Group | Group-Object dateworked
        foreach ($techDay in $techDays) {
            if ($techDay.Name) {
                # $dt = [DateTime]::ParseExact($techDay.name, "d/MM/yyyy h:mm:ss tt", [System.Globalization.CultureInfo]::CreateSpecificCulture("en-NZ")    
            
                try {
                    $dt = [DateTime]::ParseExact($techDay.name , $format, $culture)

                }
                catch {
                    write-host "Get-ATDailyTimeStats: Error occurred while parsing date for resource $($Resource.username) on Date $($techDay.name)" -ForegroundColor Red
                    write-host "there are $($techDay.group.Count) time entries for this date" -ForegroundColor Yellow
                    $techDay.Group
                    return
                }


                $result = [PSCustomObject]@{
                    resourceID                    = $Resource.id
                    Resource                      = $Resource.username
                    workDate                      = $dt
                    AODate                        = $techDay.group | select-Object OADate -First 1 | select-object -ExpandProperty OADate
                    HoursExpectedPerDay           = $Resource.dailyHrs  
                    hoursWorked                   = ($techDay.group | Measure-Object -Property hoursWorked -sum).sum
                    HrsClient                     = ($techDay.group | Measure-Object -Property HrsClient -sum).sum
                    hrsTicketBIllableNormalHrs    = ($techDay.group | Measure-Object -Property HrsClientBillableNormalHrs -sum).sum
                    hrsTicketBIllableAfterHrs     = ($techDay.group | Measure-Object -Property HrsClientBillableAfterHrs -sum).sum
                    hrsTicketNonBIllableNormalHrs = ($techDay.group | Measure-Object -Property HrsClientNonBBillableNormalHrs -sum).sum
                    hrsTicketNonBIllableAfterHrs  = ($techDay.group | Measure-Object -Property HrsClientNonBBillableAfterHrs -sum).sum
                    HrsLeave                      = ($techDay.group | Measure-Object -Property HrsLeave -sum).sum
                    HrsSick                       = ($techDay.group | Measure-Object -Property HrsSick -sum).sum
                    HrsTeaBreaks                  = ($techDay.group | Measure-Object -Property HrsTeaBreaks -sum).sum
                    HrsTraining                   = ($techDay.group | Measure-Object -Property HrsTraining -sum).sum
                    HrsInternalProd               = ($techDay.group | Measure-Object -Property HrsInternalProd -sum).sum
                    HrsInternalOther              = ($techDay.group | Measure-Object -Property HrsInternalOther -sum).sum
                    #     InternalTicketBillableNormalHrs    = ($techDay.group | Measure-Object -Property InternalTicketBillableNormalHrs -sum).sum
                    #    InternalTicketBillableAftHrs       = ($techDay.group | Measure-Object -Property InternalTicketBillableAftHrs -sum).sum
                    # InternalTicketNonBillableNormalHrs = ($techDay.group | Measure-Object -Property InternalTicketNonBillableNormalHrs -sum).sum
                    # InternalTicketNonBillableAftHrs    = ($techDay.group | Measure-Object -Property InternalTicketNonBillableAftHrs -sum).sum
                    # InternalTicketTotal                = ($techDay.group | Measure-Object -Property InternalTicket -sum).sum
                    AfterHours                    = ($techDay.group | Measure-Object -Property AfterHours -sum).sum
                    RMMTicket                     = ($techDay.group | Measure-Object -Property RMMTicket -sum).sum
                    RMMTask                       = ($techDay.group | Measure-Object -Property RMMTask -sum).sum
                }
    
                if ($Resource.DailyAvailabilities) {
                    # $DayNum = ([datetime]($Result.workDate)).DayOfWeek.value__

                    try {
                        # $DayNum = ([datetime]($Result.workDate)).DayOfWeek.value__
                        $DayNum = ($Result.workDate).DayOfWeek.value__
                        $dayweek = ($Result.workDate).DayOfWeek
                        $themonth = ($Result.workDate).ToString("MMMM")
                        #  $DayNum = ([datetime]($Result.aodate)).DayOfWeek.value__
                        Write-Host "Get-ATDailyTimeStats: Processing resource $($Resource.username) for workdate $($result.workDate) which is day number $DayNum | $dayweek | $themonth" -ForegroundColor Cyan
                    }
                    catch {
                        Write-Warning "Get-ATDailyTimeStats: Failed to get workdate on resource $($Resource.username) . Error: $_"
                        #  $result.dailyHrs = 0
                    }
                    switch ($DayNum) {
                        1 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.MondayAvailableHours } 
                        2 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.TuesdayAvailableHours }
                        3 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.WednesdayAvailableHours }
                        4 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.ThursdayAvailableHours }
                        5 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.FridayAvailableHours }
                        6 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.SaturdayAvailableHours }
                        0 { $result.HoursExpectedPerDay = $Resource.DailyAvailabilities.SUndayAvailableHours }
                        Default {}
                    }
                
                    Write-Debug "Get-ATDailyTimeStats: Day hours for $($result.Resource) on day $daynum is $($result.HoursExpectedPerDay) "
                }
                else {
                    Write-Debug "Get-ATDailyTimeStats: Day hours for $($result.Resource) were not found"
                }
                $OneResourceDates += $result
                $result
            }

            #now check for the working dates that were missing a record.

            $missingdays = $datesToCheck | Where-Object { ($_.date -ge $Resource.hireDate ) -and ($_.dateStr -notin $OneResourceDates.workDate) }
          
            [psobject[]]$MissingWorkingDays = $null

            if ($missingdays) {
                foreach ($aday in $missingdays) {
                    $DayNum = ([datetime]($aday.date)).DayOfWeek.value__
                    switch ($DayNum) {
                        1 { if ($Resource.DailyAvailabilities.MondayAvailableHours -gt 0) { $MissingWorkingDays += $aday } } 
                        2 { if ( $Resource.DailyAvailabilities.TuesdayAvailableHours -gt 0) { $MissingWorkingDays += $aday } }
                        3 { if ( $Resource.DailyAvailabilities.WednesdayAvailableHours -gt 0) { $MissingWorkingDays += $aday } }
                        4 { if ( $Resource.DailyAvailabilities.ThursdayAvailableHours -gt 0) { $MissingWorkingDays += $aday } }
                        5 { if ( $Resource.DailyAvailabilities.FridayAvailableHours -gt 0) { $MissingWorkingDays += $aday } }
                        6 { if ( $Resource.DailyAvailabilities.SaturdayAvailableHours -gt 0) { $MissingWorkingDays += $aday } }
                        0 { if ( $Resource.DailyAvailabilities.SundayAvailableHours -gt 0) { $MissingWorkingDays += $aday } }
                        Default {}
                    }
                    foreach ($aday in $MissingWorkingDays) {
                        #$Blankresult = 
                        [PSCustomObject]@{
                            resourceID                         = $Resource.id
                            Resource                           = $Resource.username
                            workDate                           = $aday.dateStr
                            HoursExpectedPerDay                = 0.0  
                            hoursWorked                        = 0.0
                            hrsTicketBIllableNormalHrs         = 0.0
                            hrsTicketBIllableAfterHrs          = 0.0
                            hrsTicketNonBIllableNormalHrs      = 0.0
                            hrsTicketNonBIllableAfterHrs       = 0.0
                            HrsLeave                           = 0.0
                            HrsSick                            = 0.0
                            HrsTeaBreaks                       = 0.0
                            HrsTraining                        = 0.0
                            HrsInternalProd                    = 0.0
                            HrsInternalOther                   = 0.0
                            InternalTicketBillableNormalHrs    = 0.0
                            InternalTicketBillableAftHrs       = 0.0
                            InternalTicketNonBillableNormalHrs = 0.0
                            InternalTicketNonBillableAftHrs    = 0.0
                            InternalTicketTotal                = 0.0
                            TicketTotal                        = 0.0
                            AfterHours                         = 0.0
                            RMMTicket                          = 0.0
                            RMMTask                            = 0.0
                        }
                    }
                }
           
            }
        }
    }
}

function Get-ATEngineers() {
    <#
    .SYNOPSIS
    Returns a list of Autotask resources (engineers/technicians) with availability data.

    .DESCRIPTION
    Queries the Autotask Resources endpoint and returns active or matching resources,
    enriched with a FullName property and a DailyAvailabilities child object from the
    ResourceDailyAvailabilities endpoint.

    Resources can be filtered by numeric ID array (ByID parameter set) or by email
    address (ByEmail parameter set). When neither is provided, all resources are returned
    excluding user-type 17 (API-only users).

    .PARAMETER id
    One or more numeric Autotask Resource IDs to retrieve. Uses the ByID parameter set.

    .PARAMETER email
    Email address of a specific resource to retrieve. Uses the ByEmail parameter set.
    Must be a valid email format; returns an error if not.

    .PARAMETER IncludeAllFields
    Switch. When set, all available resource fields are returned.
    By default a reduced field set is used: id, userName, firstName, lastName, email,
    resourceType, isActive, mobilePhone, payrollIdentifier, userType, title, hireDate.

    .PARAMETER isActive
    Switch. When set, limits results to active resources only.

    .EXAMPLE
    Get-ATEngineers

    .EXAMPLE
    Get-ATEngineers -id 29683001, 29683002

    .EXAMPLE
    Get-ATEngineers -email "jane.smith@example.com" -isActive

    .NOTES
    DailyAvailabilities is joined per-resource from v1.0/ResourceDailyAvailabilities.
    The FullName property is synthesised as "$FirstName $LastName".
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByID')]
    param (
        [Parameter(ParameterSetName = 'ByEmail', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
        [string]$email = $null,

        [Parameter(ParameterSetName = 'ByID', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
        #[int32[]]$id = @(), 
        [int32[]]$id = @(), 
        [switch]  
        $IncludeAllFields = $false,
        # [single]
        # $DeafultDailyHrs = 0,
        [switch]
        $isActive = $false
    )
    write-verbose "Polling Autotask about Resources (Engineers)"
    $includeFields = $null
    if (!$IncludeAllFields) {
        $includeFields = "id", "userName", "firstName", "LastName", "email", "resourceType", "isActive", "mobilePhone", "payrollIdentifier", "userType", "title", "hireDate"
        $IncludeFields = ('"{0}"' -f ($includeFields -join '","')).replace('""', '"')
    }
    $result = @()
    if ( $id.count -gt 0 ) {
        # $id is declared as [int32[]] so it is always an array — iterate regardless of count

        if ($id.count -eq 1) {
            $ida = $id[0]
            $result += Invoke-ATQuery -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFirstBy id -ID $ida  #-isActive:$isActive
        }
        else {
        
            $ids = $id -join ","
            Write-Verbose "Read-AutoTaskEnginners: ID's text in search is $ids "
            $result += Invoke-ATQuery -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFirstBy Nothing -SearchFurtherBy  $('{"op":"in","Field":"id","value":[' + $ids + ']}')  
        }
 

    }   
    elseif ($email) {


        if ($email -as [System.Net.Mail.MailAddress]) {
            $result += Invoke-ATQuery -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""email"",""value"":""$email""}"  -isActive:$isActive
            write-verbose "Get-ATEngineers: searched for email $email and found $($result.count) results"        
        } 
        else {
            Write-Host "Get-ATEngineers : The email address '$email' is not valid." -ForegroundColor Red
            return
        }

    }
    else {
        $result = Invoke-ATQuery -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFurtherBy '{"op":"noteq","Field":"userType","value":"17"}'  -isActive:$isActive
    }
    
    $result = $result | Select-Object * -Unique
    if ($result) {

        $result | Add-Member -NotePropertyName FullName -NotePropertyValue ""
        # $result | Add-Member -NotePropertyName dailyHrs -NotePropertyValue $DeafultDailyHrs
        $DailyAvialabilities = Invoke-ATQuery -entityName 'v1.0/ResourceDailyAvailabilities'

        foreach ($Resource in $result) {
            $resource.FullName = "$($resource.FirstName) $($resource.LastName)"
            $item = $DailyAvialabilities | Where-Object resourceID -eq $Resource.ID | Select-Object -First 1
            if ($item) {
                Write-Debug "Read-Engineers: found availabilities for $($resource.username) :$($item -join (',')) of availabilities"
                $resource | Add-Member  -NotePropertyName 'DailyAvailabilities' -NotePropertyValue $Item
                #  write-debug "Read-Engineers: Monday availability for $($resource.username) is $($resource.DailyAvailabilities.MondayAvailableHours)"
                #$resource.DeafultDailyHrs = $item.sundayAvailableHours + $item.MondayAvailableHours+ $item.TuesdayAvailableHours+ $item.WednesdayAvailableHours+ $item.ThursdayAvailableHours+ $item.FRidayAvailableHours+ $item.SaturdayAvailableHours
            }
            #else { $Resource.DeafultDailyHrs = $DeafultDailyHrs}

        }


        # ($result | Where-Object userName -eq "rogelio.vera").dailyHrs = 4
        write-verbose "DONE Polling Autotask about Resources (Engineers)" #-ForegroundColor Green
        return $result
    }
}


# function Read-AutotaskQueues() {
#     [CmdletBinding()]
#     param (
#         # [Parameter()]
#         # [TypeName]
#         # $ParameterName
#     )

#     $result = Invoke-ATQuery -entityName 'v1.0/Resources' #-includeFields $includeFields -SearchFurtherBy '{"op":"noteq","Field":"userType","value":"17"}'  -isActive:$isActive
 
# }





function Get-ATEngineerToReport() {
    <#
    .SYNOPSIS
    Retrieves the Autotask Resource ID (ATResourceID) saved in the local login file.

    .DESCRIPTION
    Reads the kiss-atapi login JSON file and returns the ATResourceID integer that was
    previously stored by Set-ATEngineerToReport.  This ID identifies which
    engineer's time entries and data are used as the default when running reports
    with the -ForMeOnly switch.

    If no ATResourceID is found, a warning is displayed and nothing is returned.

    .EXAMPLE
    $myID = Get-ATEngineerToReport
    Get-ATTimeEntries -ForResourceID $myID -LastxMonths 1

    .NOTES
    Run Set-ATEngineerToReport first to persist the engineer ID.
    Requires that Set-ATLogin has already been run so the login file exists.
    #>
    [CmdletBinding()]
    param (
        
    )

    if (test-path -path "$kissATAPIpath\$kissATAPIfile" ) {
        $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
        if ($jsn) {
            #write-host "there was a prexisting saved login of $jsn"
            $r = $jsn  | ConvertFrom-Json 
            if ($r | Get-Member -Name "ATResourceID" -ErrorAction SilentlyContinue) {
                #if (-not $r.PSObject.Properties.Match("ATResourceID")) {
                # Create it if missing
                # Update it if it exists
                write-verbose "Found a saved ATResourceID of $($r.ATResourceID) in the saved login, returning this value." #-ForegroundColor Green
                #$r.ATResourceID
                return  [int] $r.ATResourceID
            }
            else {
                # $r
                write-Host "No ATResourceID found in the saved login, you need to run Set-AutoTaskEngineerToReport to set an engineer to report on." -ForegroundColor Yellow

            }
                        
               
        }

    }

    Write-Host "Get: No global Autotask API or ATResourceID is currently set. PLease run Set-ATLogin and  Set-AutoTaskEngineerToReport to set an engineer to report on." -ForegroundColor Yellow
       

}

function Get-ATLastQuickNote() {
    <#
    .SYNOPSIS
    Returns the most recent Quick Note (action type 5) for each company in Autotask.

    .DESCRIPTION
    Queries the Autotask CompanyNotes endpoint for all notes of action type 5 (Quick Note)
    created after 2018-01-01.  Results are grouped by company, and the single most recently
    modified note per company is returned, along with a NoteCount property showing how many
    quick notes that company has in total.

    .EXAMPLE
    $latestNotes = Get-ATLastQuickNote
    $latestNotes | Select-Object CompanyID, NoteCount, lastModifiedDate, Note

    .NOTES
    The date filter (after 2018-01-01) is hard-coded in the URL.  No parameters are exposed.
    #>
    [CmdletBinding()]
    param (  )
    $rc = Invoke-ATQuery -url "https://webservices6.autotask.net/ATServicesRest/V1.0/CompanyNotes/query?search={""IncludeFields"":[""id"",""CompanyID"",""lastModifiedDate"",""Note""],""filter"":[{""op"":""and"",""items"":[{""op"":""eq"",""field"":""actionType"",""value"":""5""},{""op"":""gt"",""field"":""lastModifiedDate"",""value"":""2018-01-01T00:00:00.00Z""}]}]}"
    $rc = $rc | Group-Object companyID

    foreach ($item in $rc) {
        $item[0].Group | Add-Member -NotePropertyName NoteCount -NotePropertyValue $item.Count
        ($item[0].Group | Sort-Object -Descending lastModifiedDate)[0]
    }

}


function Get-ATMostRecentCompanyTicket() {
    <#
    .SYNOPSIS
    Returns the most recently completed ticket for each company that has an assigned resource.

    .DESCRIPTION
    Polls Autotask for all completed tickets since 2020-01-01 that have a resource assigned,
    then groups them by CompanyID and returns the single most recently completed ticket per company.
    Useful for gauging the last time any work was done for each customer.

    .EXAMPLE
    $latestTickets = Get-ATMostRecentCompanyTicket
    $latestTickets | Select-Object companyID, completedDate, title | Sort-Object completedDate -Descending

    .NOTES
    The start date (2020-01-01) is hard-coded.
    Only tickets with a completedDate and an assigned resource are included.
    #>
    [CmdletBinding()]
    param (  )
    $rc = Get-ATTickets -LastActionFromDate "2020-01-01T00:00:00" -Verbose -DontexpandticketInformation -whereResourceAssigned -includeFields ("companyID", "completedDate", "id", "title", "createDate") -ExcludeNonComplete | Group-Object companyID
    #$rc.psobject.properties.remove('userDefinedFields')
    foreach ($item in $rc) {
        ($item[0].Group | Sort-Object -Descending completedDate)[0]
    }
}


function Get-ATRoles() {
    <#
    .SYNOPSIS
    Returns all role records from Autotask.

    .DESCRIPTION
    Retrieves every role defined in Autotask (v1.0/Roles).
    Roles are assigned to resources and may appear on time entries via the roleID field.
    Use this to translate a numeric roleID into a human-readable role name.

    .EXAMPLE
    $roles = Get-ATRoles
    $roles | Select-Object id, name

    .NOTES
    No parameters. Returns all roles without pagination (typically a small dataset).
    #>
    [CmdletBinding()]
    param (
    )
    Write-verbose "Polling Autotask for Role Codes" #-ForegroundColor Green"
    $includeFields = ('id', 'name')
    Invoke-ATQuery -entityName 'v1.0/Roles' -includeFields $includeFields

}

function Get-ATTicketFieldInfo {
    <#
    .SYNOPSIS
    Returns picklist metadata for key Autotask ticket fields.

    .DESCRIPTION
    Queries the Autotask Tickets entity-information endpoint to retrieve the available
    picklist values for the following fields:
      queueID, status, issueType, monitorTypeID, TicketCategory, ticketType

    Returns a single PSCustomObject whose properties each contain the picklist array
    for that field. Each picklist item has value, label, and isActive properties.

    Useful for translating raw numeric field values (e.g. status = 5) into human-readable
    labels without making additional API calls per ticket.

    .PARAMETER ExportCSV
    Switch. Reserved for future use — not yet implemented.

    .EXAMPLE
    $fields = Get-ATTicketFieldInfo
    $fields.status | Where-Object isActive | Select-Object value, label

    .EXAMPLE
    # Translate a status code
    ($fields.status | Where-Object value -eq 5).label

    .NOTES
    The endpoint used is v1.0/Tickets/entityInformation/fields.
    Only a subset of fields is exposed — other fields (e.g. priority) can be added
    by extending the returned PSCustomObject.
    #>
    [CmdletBinding()]
    param (
        [switch]$ExportCSV
    )

    Write-Host "Read-TicketInformation Polling Autotask for TicketInformation queues, status etc. values "
    $fields = (Invoke-ATQuery -UrlFixedSuffix v1.0//Tickets/entityInformation/fields).fields #(name,picklistvalues[value,label,isactive)

    [PSCustomObject]@{
        queueID        = ($fields | where-object name -eq "queueID" | Select-Object  * -First 1).picklistValues
        status         = ($fields | where-object name -eq "status" | Select-Object  * -First 1).picklistValues 
        issueType      = ($fields | where-object name -eq "issueType" | Select-Object  * -First 1).picklistValues 
        monitorTypeID  = ($fields | where-object name -eq "monitorTypeID" | Select-Object  * -First 1).picklistValues 
        TicketCategory = ($fields | where-object name -eq "TicketCategory" | Select-Object  * -First 1).picklistValues 
        ticketType     = ($fields | where-object name -eq "ticketType" | Select-Object  * -First 1).picklistValues 
    }

    Write-Host "DONE-Read-TicketInformation Polling Autotask for Read-TicketInformation queues, status etc. values" -ForegroundColor Green
}

function Get-ATTickets {
    <#
    .SYNOPSIS
    Retrieves tickets from Autotask with flexible filtering options.

    .DESCRIPTION
    Polls the Autotask Tickets endpoint and returns a collection of ticket objects.
    Supports filtering by ticket ID(s), company ID(s), ticket number(s), title content,
    date of last activity, assigned resource, and completion status.

    By default, each ticket is enriched with human-readable QueueName, StatusName, and
    ResourceName fields by looking up the ticket field metadata and the engineers list.
    Use -DontexpandticketInformation to skip this enrichment for faster results.

    Results are automatically de-duplicated on TicketNumber.

    .PARAMETER ids
    One or more specific Autotask ticket IDs to retrieve.

    .PARAMETER CompanyIDs
    One or more Autotask company IDs. Only tickets belonging to these companies are returned.

    .PARAMETER TicketNumbers
    One or more Autotask ticket numbers (e.g. 'T20240101.0001').

    .PARAMETER LastActionFromDate
    Return only tickets whose lastActivityDate is on or after this date.
    Defaults to 60 days ago.

    .PARAMETER TitleContains
    Filter tickets whose title contains this string.

    .PARAMETER TitleBeginsWith
    Filter tickets whose title begins with this string (e.g. 'RMM').

    .PARAMETER includeFields
    Array of field names to include in the response. If omitted a standard default set is used.

    .PARAMETER ReturnAllFields
    Switch. When set, all available ticket fields are returned and includeFields is ignored.

    .PARAMETER IncludeAllNonComplete
    Switch. When set, also returns all open (non-completed) tickets that have an assigned resource,
    regardless of last activity date.

    .PARAMETER ExcludeNonComplete
    Switch. When set, only returns tickets that have a completedDate.
    .PARAMETER ExcludeComplete
    Switch. When set, only returns tickets that not have a completedDate.

    .PARAMETER DontexpandticketInformation
    Switch. Skips the secondary lookup of queue/status/resource names. Faster but returns raw IDs only.

    .PARAMETER whereResourceAssigned
    Switch. Limits results to tickets that have an assigned resource.

    .PARAMETER LastxWeeks
    If supplied, overrides LastActionFromDate and filters for tickets active within the last N days.

    .PARAMETER loopCount
    Maximum number of recursive pages to retrieve (500 records per page). Default 40 (20,000 tickets).

    .PARAMETER DoSearchBy
    Provides a fully custom Autotask filter JSON string, bypassing all other filter logic.

    .PARAMETER GetCompanyNames
    Switch. Enriches results with CompanyName and CompanyClassification by calling Get-ATCompanies
    for each unique CompanyID. Can be slow with large result sets.

    .PARAMETER ForCalendar
    Switch. Returns a minimal field set optimised for calendar display and skips ticket information expansion.

    .EXAMPLE
    Get-ATTickets -LastActionFromDate (Get-Date).AddDays(-30)

    .EXAMPLE
    Get-ATTickets -ids 12345, 67890

    .EXAMPLE
    Get-ATTickets -TitleBeginsWith "RMM" -LastActionFromDate "2024-01-01"

    .EXAMPLE
    Get-ATTickets -IncludeAllNonComplete -whereResourceAssigned -GetCompanyNames

    .NOTES
    Results are de-duplicated on TicketNumber using CheckDuplicatesOf inside Invoke-ATQuery.
    Date strings passed to LastActionFromDate should be parseable by PowerShell's [DateTime] cast.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByFilter')]
    param (
        [Parameter(ParameterSetName = 'ByID', Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [int[]]$ids, #= @(),
        [Parameter(ParameterSetName = 'ByCompany', Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [int[]]$CompanyIDs ,#= @(),
        [Parameter(ParameterSetName = 'ByNumber', Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$TicketNumbers,# = @(),
        [DateTime]
        [Nullable[DateTime]]$LastActionFromDate , # (Get-date).AddDays(-60),
        [Nullable[int]]$LastxWeeks = -1,
        [string]$TitleContains,
        [string]$TitleBeginsWith,
        [string[]]$includeFields = $null,
        [switch]$ReturnAllFields = $false,
        [switch]$IncludeAllNonComplete = $null,
        [switch]$ExcludeNonComplete = $false,
        [switch]$ExcludeComplete = $false,
       
        [switch]$DontexpandticketInformation = $false,
        [switch]$whereResourceAssigned,
        [int] $loopCount = 40,
        [string]$DoSearchBy = $null,
        [switch]$GetCompanyNames = $false,
        [switch]$ForCalendar = $false
    )
    write-verbose "Get-ATTickets: polling autotask for ticket information"
    $ticketinfo = $null
    $extraSearchs = @()



    if ($ForCalendar -eq $true) {
        
        write-verbose "Get-ATTickets: ForCalendar is true, so we are only getting a subset of fields to make it faster for calendar use"
        $includeFields = ('id', 'TicketNumber', 'CompanyID', 'status', 'BillingCodeID', 'tickettype', 'title')
        $GetCompanyNames = $true
        $DontexpandticketInformation = $true
    }
    # $searchby =$searchby -replace ' ',''
    if ($ReturnAllFields -eq $true) { $includeFields = $null }
    elseif (!$includeFields) {
        $includeFields = ('id', 'TicketNumber', 'CompanyID', 'completedDate', 'createDate', 'BillingCodeID', 'firstResponseDateTime', 'lastActivityDate', 'status', 'tickettype', 'completedDate', 'title', 'assignedResourceID', 'queueid','estimatedHours')
    }
    
    if (($null -ne $Lastweeks) -and ($LastxWeeks -ge 0)) {
        # Get-ATWeekStart returns midnight local time — do NOT call ToUniversalTime() here;
        # the lastActivityDate field is compared in local time by the Autotask API.
        $datesFrom = Get-ATWeekStart -LastXWeeks $LastxWeeks
        $LastActionFromDateStr = $datesFrom.ToString("yyyy-MM-ddTHH:mm:ss")
        $extraSearchs += '{"op":"gte","Field":"lastActivityDate","value":"' + $LastActionFromDateStr + '"}'
    }
    IF ($LastActionFromDate) {
        $LastActionFromDateStr = $LastActionFromDate.ToString("yyyy-MM-ddTHH:mm:ss")
        $extraSearchs += '{"op":"gte","Field":"lastActivityDate","value":"' + $LastActionFromDateStr + '"}' #+ ',' + $i 
    }   
    if ($IncludeAllNonComplete -eq $true) {
        #search for tickets that have been assigned but not yet completed
        $extraSearchs += '{"op":"and","items":[{"op":"notExist","Field":"completedDate"},{"op":"Exist","Field":"assignedResourceID"}]}'
    }
    if ($TitleContains) {
        #AND search for title contains
        $extraSearchs += '{"op":"contains","Field":"title","value":"' + $TitleContains + '"}'
    }
    if ($TitleBeginsWith) {
        #AND search for title begins with
        $extraSearchs += '{"op":"beginsWith","Field":"title","value":"' + $TitleBeginsWith + '"}'
    }
    if ($whereResourceAssigned -eq $true) {
        #AND search for assigned resource exists (ie tickets that have an assigned resource, which usually means they are being worked on)
        $extraSearchs += '{"op":"Exist","Field":"assignedResourceID"}' #+ ',' + $i
    }
    if ($ExcludeNonComplete -eq $true) {
        #and search for completed date exists (ie only include tickets that have a completed date, which means they are completed)
        $extraSearchs += '{"op":"Exist","Field":"completedDate"}' #+ ',' + $i
    }
    if ($ExcludeComplete -eq $true) {
        #and search for completed date exists (ie only include tickets that have a completed date, which means they are completed)
        $extraSearchs += '{"op":"notExist","Field":"completedDate"}' #+ ',' + $i
    }

    #------------------------------
    $items = @()

    if ($ids.count -gt 0) {
        #byID or by ID set
        if ($ids.count -eq 1) {
            $ida = $ids[0]
            write-verbose "Get-ATCompanies - for a ID $ida"
            $items = Invoke-ATQuery -entityName 'v1.0/Tickets' -id $ida -SearchFirstBy id -includeFields $includeFields
        }
        else {
            write-verbose "Get-ATTickets - for a ID SET $ids "
            $chunkSize = 100
            #$rc = @()
            for ($i = 0; $i -lt $ids.Count; $i += $chunkSize) {
                $chunk = [int[]]$ids[$i..([Math]::Min($i + $chunkSize - 1, $ids.Count - 1))]
                IF ($chunk) {
                    [string]$cci = $chunk -join ','
                    Write-verbose "Get-ATTickets ids searched for are $cci"
                    if ($extraSearchs) {
                        $thisSearch = '{"op":"and","items":[' + ($extraSearchs -join ",") + ',{"op":"in","Field":"id","value":[' + $cci + ']}]}'
                    }
                    else {
                        $thisSearch = '{"op":"in","Field":"id","value":[' + $cci + ']}'
                    }
                    $items += Invoke-ATQuery -entityName 'v1.0/Tickets' -includeFields $includeFields  -SearchFurtherBy $thisSearch -LoopCount 40  -CheckDuplicatesOf "TicketNumber"
                    # return $items
                }                    
            }
        }
        # return $items
    } 

    elseif ( $TicketNumbers.count -gt 0 ) {
        #by ticket number or set of ticket numbers
        $chunkSize = 100
        Write-Host  "Reading tickets via ticket number"
        for ($i = 0; $i -lt $TicketNumbers.Count; $i += $chunkSize) {
            $chunk = [string[]]$TicketNumbers[$i..([Math]::Min($i + $chunkSize - 1, $TicketNumbers.Count - 1))]
            IF ($chunk) {
                [string]$cci = $chunk -join '","'
                Write-host "Get-ATTickets TicketNumbers searched for are $cci"
                # $searchby = '{"op":"in","Field":"TicketNumber","value":["' + $cci + '"]}'

                if ($extraSearchs) {
                    $thisSearch = '{"op":"and","items":[' + ($extraSearchs -join ",") + ',{"op":"in","Field":"TicketNumber","value":["' + $cci + '"]}]}'
                }
                else {
                    $thisSearch = '{"op":"in","Field":"TicketNumber","value":["' + $cci + '"]}'
                }
                $items += Invoke-ATQuery -entityName 'v1.0/Tickets' -includeFields $includeFields  -SearchFurtherBy $thisSearch -LoopCount 40  -CheckDuplicatesOf "TicketNumber"
                #   return $items
            }                    
        }
    }

    elseif ($companyIDs.count -gt 0) {
        #by company ID or set of company IDs
        write-host "Get-ATTickets - that belong to specific company(s) $ids "
        $chunkSize = 100
        #$rc = @()
        for ($i = 0; $i -lt $companyIDs.Count; $i += $chunkSize) {
            $chunk = [int[]]$companyIDs[$i..([Math]::Min($i + $chunkSize - 1, $companyIDs.Count - 1))]
            IF ($chunk) {
                [string]$cci = $chunk -join '","'
                Write-host "Get-ATTickets belonging to Companyids searched for are $cci"
                if ($extraSearchs) {
                    $thisSearch = '{"op":"and","items":[' + ($extraSearchs -join ",") + ',{"op":"in","Field":"companyID","value":["' + $cci + '"]}]}'
                }
                else {
                    $thisSearch = '{"op":"in","Field":"companyID","value":["' + $cci + '"]}'
                }
                $items += Invoke-ATQuery -entityName 'v1.0/Tickets' -includeFields $includeFields  -SearchFirstBy nothing -SearchFurtherBy $thisSearch -LoopCount 40  -CheckDuplicatesOf "TicketNumber"
                #  return $items
            }
        }
        # return items
    }
    else {
        if ($extraSearchs.count -gt 0) {
            Write-Host "Get-ATTickets that have been had a last action in the last $lastxweeks weeks"
            $thisSearch = '{"op":"and","items":[' + ($extraSearchs -join ",") + ']}'
            $items += Invoke-ATQuery -entityName 'v1.0/Tickets' -includeFields $includeFields  -SearchFirstBy nothing -SearchFurtherBy $thisSearch -LoopCount 40  -CheckDuplicatesOf "TicketNumber"

        }
    }

    
    # return $items

  
    #write-host $i
    write-verbose "Get-ATTickets: search by : $searchby"
    #  $items = Invoke-ATQuery -entityName 'v1.0/Tickets' -includeFields $includeFields  -SearchFurtherBy $searchby -SearchFirstBy Nothing -LoopCount $loopCount -CheckDuplicatesOf "TicketNumber"
    Write-Verbose "Get-ATTickets: finished polling autotask for ticket information $($items.Count) , now processing the results"
    #$x = Read-Host "hhelo"
    #return $items
    if ($items) {
        # return items
        #if ($GetCompanyNames -or ($ForCalendar -eq $true)) {
        if ($GetCompanyNames ) {
            Write-Verbose "Get-ATTickets: Getting companny info about each ticket"
            $CompaniesToGet = $items | Where-Object companyID | Select-Object -Unique -ExpandProperty CompanyID
            # return $CompaniesToGet
            $Companys = Get-ATCompanies -id $CompaniesToGet -DontExpandChildIDFields
            # return $Companys
            $items | Group-Object CompanyID | ForEach-Object {
                $companyID = $_.Name
                #$companyName = (Get-ATCompanies -id $companyID).CompanyName
                $company = $Companys | Where-Object { $_.id -eq $companyID } | Select-Object  -First 1

                foreach ($ticket in $_.Group) {
                    $ticket | Add-Member -NotePropertyName CompanyName -NotePropertyValue $company.CompanyName -Force
                    $ticket | Add-Member -NotePropertyName CompanyClassification -NotePropertyValue $company.Classification -Force
                    if ($company.Classification -eq $ATInternalClasificationCode) {
                        #  $ticket | Add-Member -NotePropertyName billingCodeID -NotePropertyValue $ATnonBillableCodes[0] -Force
                        $ticket | Add-Member -NotePropertyName CompanyType -NotePropertyValue "Internal-Customer" -Force
                        $ticket | Add-Member -NotePropertyName CompanyIsInternal -NotePropertyValue $true -Force
                    }
                    else {
                        $ticket | Add-Member -NotePropertyName CompanyIsInternal -NotePropertyValue $false -Force

                    }
                }

                # $_.Group | ForEach-Object { $_ | Add-Member -NotePropertyName CompanyName -NotePropertyValue $company.CompanyName -Force }
                # $_.Group | ForEach-Object { $_ | Add-Member -NotePropertyName CompanyClassification -NotePropertyValue $company.Classification -Force }
                # if ($company.Classification -eq $ATInternalClasificationCode) {
                #     #  $_.Group | ForEach-Object { $_ | Add-Member -NotePropertyName billingCodeID -NotePropertyValue $ATnonBillableCodes[0] -Force }
                #     $_.Group | ForEach-Object { $_ | Add-Member -NotePropertyName CompanyType -NotePropertyValue "Internal-Customer" -Force }
                #     $_.Group | ForEach-Object { $_ | Add-Member -NotePropertyName CompanyIsInternal -NotePropertyValue $true -Force }
                # }
                # else {
                #     $_.Group | ForEach-Object { $_ | Add-Member -NotePropertyName CompanyIsInternal -NotePropertyValue $false -Force }

                # }
            }            

            Write-Verbose "Get-ATTickets: added company names to the ticket $($item.ticketNumber) with companyID $($item.CompanyID) and company name $($item.CompanyName)"
        }

           
    
        if (!($DontexpandticketInformation)) {
            Write-Verbose "Get-ATTickets: Getting queue and status messages for each ticket"
            $ticketinfo = Get-ATTicketFieldInfo

            $items | Add-Member -NotePropertyName QueueName -NotePropertyValue "" -Force
            $items | Add-Member -NotePropertyName StatusName -NotePropertyValue "" -Force
            $items | Add-Member -NotePropertyName ResourceName -NotePropertyValue "" -Force
            # $items |Add-Member -NotePropertyName Company -NotePropertyValue "" -Force
            
            $Resources = Get-ATEngineers
            foreach ($titem in $items) {
                $titem | Add-Member -NotePropertyName QueueName -NotePropertyValue "$((($ticketinfo.queueID) | Where-Object value -eq $titem.queueID | Select-Object label -first 1).label)" -Force
                $titem | Add-Member -NotePropertyName StatusName -NotePropertyValue "$((($ticketinfo.status) | Where-Object value -eq $titem.status | Select-Object label -first 1).label)" -Force

                # $titem.QueueName = (($ticketinfo.queueID) | Where-Object value -eq $titem.queueID | Select-Object label -first 1).label
                # $titem.StatusName = (($ticketinfo.status) | Where-Object value -eq $titem.status | Select-Object label -first 1).label
                if ($titem.assignedResourceID) {
                    # $titem.ResourceName = ($Resources  | Where-Object id -eq $titem.assignedResourceID | Select-Object  -first 1).userName 
                    $titem | Add-Member -NotePropertyName ResourceName -NotePropertyValue "$(($Resources  | Where-Object id -eq $titem.assignedResourceID | Select-Object  -first 1).userName )" -Force

                }
                else {
                    $titem | Add-Member -NotePropertyName ResourceName -NotePropertyValue "" -Force
                }
            } 
            Convert-ObjArrayDateTimesToSearchableStrings $items 
        }
    }
    $items
    write-verbose "DONE -Get-ATTickets: have polled autotask for ticket information"# -ForegroundColor Green
}

function Get-ATTimeEntries() {
    <#
    .SYNOPSIS
    Polls Autotask for time entries and returns an enriched, classified array.

    .DESCRIPTION
    Retrieves timesheet entries from Autotask and annotates each entry with calculated
    fields covering billable vs non-billable hours, after-hours, leave, internal work,
    RMM activity, and work-type classification (kissWorkType).

    Date range is controlled by one of: LastxMonths, LastXWeeks, FromDateLocal, or
    FromDateUTC. Leave entries are excluded by default; pass -includeLeave to retain them.

    Only ticket-linked entries are returned when -ForCalendar is used.
    When a ticket number is supplied via -ForTicketNumber, it is resolved to an ID first
    via Get-ATTickets, then entries are filtered to that ticket only.

    WARNING: Do not rely on the Autotask field hoursToBill for accurate billing totals -
    it can exceed hoursWorked in some cases. Use hoursBillable (added by this function)
    instead.

    .PARAMETER LastxMonths
    How many calendar months back to start pulling time entries from. Default 0 (disabled).

    .PARAMETER LastXWeeks
    How many weeks back to start, counting from the previous Sunday. Default 0 (disabled).

    .PARAMETER FromDateLocal
    Explicit local DateTime to use as the start of the retrieval window.

    .PARAMETER FromDateUTC
    Explicit UTC DateTime to use as the start of the retrieval window.
    Converted to local time internally before building the API filter.

    .PARAMETER ForMeOnly
    Switch. Filters entries to the engineer saved via Set-ATEngineerToReport.
    Throws if no engineer has been saved.

    .PARAMETER ForResouerceID
    Numeric Autotask Resource ID to filter entries to a specific engineer.

    .PARAMETER ForTicketID
    Filter entries to a single ticket by its numeric Autotask ticket ID.

    .PARAMETER ForTicketNumber
    Filter entries to a single ticket by its ticket number (e.g. 'T20260101.0001').
    Resolved to a ticket ID internally via Get-ATTickets -TicketNumbers.

    .PARAMETER ReturnRaw
    Switch. Returns the raw API response without any enrichment or classification.

    .PARAMETER ForCalendar
    Switch. Returns a minimal field set sorted by endDateTime, suitable for calendar display.
    Only entries that have a TicketID are included.

    .PARAMETER DisplayCompanySummary
    Switch. After processing, prints a per-company and per-work-type hours summary to the host.

    .PARAMETER IncludeSummaryNotes
    Switch. Includes the summaryNotes field in the returned entries.

    .PARAMETER includeLeave
    Switch. Includes leave and sick-leave entries (excluded by default).

    .PARAMETER ReturnAllFields
    Switch. Returns all fields from the API rather than the default reduced field set.

    .PARAMETER afterHrsBillingCodes
    Billing code IDs to classify as after-hours work. Default: 29683343, 29737351.

    .PARAMETER LeaveCodes
    Billing code IDs to classify as leave. Default: values from $CONSTATLeaveCodes.

    .PARAMETER SickCodes
    Billing code IDs to classify as sick leave. Default: values from $CONSTATSickCodes.

    .PARAMETER teabreakCodes
    Billing code IDs to classify as tea breaks. Default: 91209.

    .PARAMETER TrainingCodes
    Billing code IDs to classify as training. Default: 29683344.

    .PARAMETER ProductiveCodes
    Billing code IDs to classify as internal-productive work.

    .PARAMETER RMMCode
    Billing code ID used to identify RMM task entries. Default: 29712660.

    .PARAMETER ATInternalClasificationCode
    Company classification numeric value that identifies an internal (non-client) company.
    Default: value of $CONSTInternalClasificationCode (200).

    .EXAMPLE
    $i = Get-ATTimeEntries -LastxMonths 3

    .EXAMPLE
    Get-ATTimeEntries -LastXWeeks 1 -ForMeOnly -DisplayCompanySummary

    .EXAMPLE
    Get-ATTimeEntries -ForTicketNumber 'T20260101.0001' -IncludeSummaryNotes

    .NOTES
    Calls Get-ATTickets internally when -ForTicketNumber is used or when ticket detail
    enrichment is needed (companyID, classification, etc.).
    Results are not de-duplicated; each Autotask time entry appears once.
    #>
    [CmdletBinding()]
    param (
        # Parameter help The number of months earlier than now, from which to start pulling the time sheeting data from
        [Parameter()]
        [int]
        $LastxMonths = 0,
        # Parameter help description
        [Parameter()]
        [int]$LastXWeeks = -1,
        [Nullable[DateTime]]$FromDateLocal = $null,
        [Nullable[DateTime]]$FromDateUTC = $null,
        [Nullable[DateTime]]$UntilDateLocal = $null,
        [switch]$ForMeOnly = $false,
        [Parameter( ValueFromPipelineByPropertyName = $true)]
        [Nullable[int]]$ForResourceID = $null,
        [Parameter( ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [NUllable[int]]$ForTicketID = $null,
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ForTicketNumber = $null,
        # [Parameter(  ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        # [string]$ForCustomerID= $null,

        [switch]$ReturnRaw = $false,
        [switch]$ForCalendar = $false,
        #  [switch]$ForCalendarLiteView = $false,
        [switch]$DisplayCompanySummary = $false,
        [switch]$IncludeSummaryNotes = $false,
        #[switch]$IncludeBillingDetails = $false,
        # [switch]$includeTicketDetails = $false,
        # [switch]$includeEngineerDetails = $false,
        [switch]$includeLeave = $false,
        [switch]$ReturnAllFields = $false,
        [ValidateNotNullOrEmpty()]
        #ticket codes
        [int[]]
        $afterHrsBillingCodes = @(29683343, 29737351),
        # Parameter help these BillingCodes are such as Sick Leave or Holidays and thus shouldn't be measured during productivity %
        # [int[]]$ATnonBillableCodes = @(29682861), #Non Billable Support
        # [ValidateNotNullOrEmpty()]
        #------------internal codes
        #[Parameter(AttributeValues)]
        [ValidateNotNullOrEmpty()]
        [int[]]
        $LeaveCodes = $CONSTATLeaveCodes,  #@(91206, 29718729),
        [ValidateNotNullOrEmpty()]
        # Parameter help these BillingCodes are such as Sick Leave or Holidays and thus shouldn't be measured during productivity %
        #[Parameter(AttributeValues)]
        [int[]]
        $SickCodes = $CONSTATSickCodes,
        [ValidateNotNullOrEmpty()] 
        [int[]]
        $teabreakCodes = @(91209),
        [ValidateNotNullOrEmpty()]
        [int[]]
        $TrainingCodes = @(29683344), #, training
        [ValidateNotNullOrEmpty()]
        [int[]]
        $ProductiveCodes = @(29711172, 29712660, 29713657, 29737360, 29718730, 29737360), #Second Level Support, RMM, presales, research, renewals

        [int]
        $RMMCode = 29712660,
        # [ValidateNotNullOrEmpty()]
        [int]$ATInternalClasificationCode = $script:CONSTInternalClasificationCode






        
    )

    write-verbose "Get-ATTimeEntries: Polling AutoTask for TimeEntries, and formating the results"
    $CURRENTDATE = GET-DATE -Hour 0 -Minute 0 -Second 0
    $CurrentEndDateStr = (Get-Date -HOUR 23 -Minute 59 -Second 59).ToString("yyyy-MM-ddTHH:mm:ss")
    $Monthstart = $null
    # $searchbyEndDate = $false
    If ($null -ne $FromDateLocal) {
        $Monthstart = $FromDateLocal
        Write-Verbose "Get-ATTimeEntries: Filtering time entries by local dateWorked >= $Monthstart"
    }
    elseif ($null -ne $FromDateUTC) {
        $Monthstart = $FromDateUTC.ToLocalTime()
        #   $searchbyEndDate = $true
        write-Verbose "Get-ATTimeEntries: Filtering time entries by local dateWorked >= $Monthstart (which is $FromDateUTC in UTC)"
    }
    elseif ($LastXWeeks -ge 0) {
        # #calculates the start of the nth week ago, where n is $LastXWeeks, and the week starts on Sunday 
        # $wday = [int](Get-Date).DayOfWeek
        # if ($wday -eq 0) { $wday = 7 } # convert Sunday from 0 to 7 for easier calculations
        # $Monthstart = $CURRENTDATE.AddDays(-((7 * ($lastXWeeks - 1)) + $wday))
        #new way to get the Sunday of n weeks ago, where n is $LastXWeeks, and the week starts on Sunday
        $Monthstart = Get-ATWeekStart -LastXWeeks $LastXWeeks
        Write-Verbose "Get-ATTimeEntries: Filtering time entries by local dateWorked >= $Monthstart (which is the Sunday of $LastXWeeks ago)"
    }
    elseif ($LastxMonths -gt 0) {
        $Monthstart = $CURRENTDATE.AddMonths(-$LastxMonths)
        write-Verbose "Get-ATTimeEntries: Filtering time entries by local dateWorked >= $Monthstart (which is $LastxMonths months ago)"
    }  
  


    #[DateTimeOffset]::Now.Offset - how to adjust for local timezone offset when comparing to API UTC times
    #$utc.ToLocalTime()


    if ($IncludeSummaryNotes -eq $true) {
        $includefields = "id", "startDateTime", "endDateTime", "billingCodeID", "roleID", "taskID", "ticketID", "timeEntryType", "resourceID", "isNonBillable", "hoursWorked", "dateWorked", "summaryNotes"
    }
    else {
        $includefields = "id", "startDateTime", "endDateTime", "billingCodeID", "roleID", "taskID", "ticketID", "timeEntryType", "resourceID", "isNonBillable", "hoursWorked", "dateWorked"
    }

    $filters = @()

    if ($null -ne $BeforeDate){
     #   $UntilDateUTCStr = $UntilDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $beforeDateLocalStr = $BeforeDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
   #     $filters += '{"op":"notIn","Field":"billingCodeID","value":[' + $(($LeaveCodes + $SickCodes) -join ",") + ']}'
       $filters += '{"op":"lt","Field":"dateWorked","value":"' + $beforeDateLocalStr + '"}'

    }

    if ($Monthstart) {
        $MonthStartLocalStr = $Monthstart.ToString("yyyy-MM-ddTHH:mm:ss")
        $MonthStartSTr = $Monthstart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        write-host "return only entries of (UTC)endDateTime >=  $MonthStartSTr  OR (local)WorkDate >=  $MonthStartLocalSTr" -ForegroundColor Green
        $filters += '{"op":"or","items":[' + '{"op":"gte","Field":"dateWorked","value":"' + $MonthStartLocalSTr + '"},' + '{"op":"gte","Field":"endDateTime","value":"' + $MonthStartSTr + '"}' + ']}'
    }

    if ($UntilDateLocal) {
        $UntilDateLocalStr = $UntilDateLocal.ToString("yyyy-MM-ddTHH:mm:ss")
        write-host "return only entries where WorkDate <=  $UntilDateLocalStr" -ForegroundColor Green
        $filters += '{"op":"lte","Field":"dateWorked","value":"' + $UntilDateLocalStr + '"}'
    }


    if ($includeLeave -ne $true) {
        #exclude leave related time entries by default, unless the user specifically wants to include them by setting includeLeave to true, in which case we will not filter them out
        $filters += '{"op":"notIn","Field":"billingCodeID","value":[' + $(($LeaveCodes + $SickCodes) -join ",") + ']}'
    }
   
    # #make sure we do not retrive entries with a work date in the future, as these are likely to be Leave  or test entries or misdated entries that would skew our reporting, and we want to focus on past and current entries for accurate reporting
    $filters += '{"op":"lte","Field":"dateWorked","value":"' + $CurrentEndDateStr + '"}'

    $ticketToGet = $null 
    If ($ForTicketID) {
        $ticketToGet = $ForTicketID
    }
    elseif ($ForTicketNumber) {

        $i = Get-ATTickets -TicketNumbers $ForTicketNumber -ForCalendar
        if ($i) {
            $ticketToGet = $i.ID
        }
        else {
            Write-Host "No ticket found with number $ForTicketNumber" -ForegroundColor Yellow
            return
        }
    }
    
    if ($ticketToGet) {
        $filters = @()
        write-verbose "Filtering by ticketID: $ticketToGet"
        $filters += '{"op":"eq","Field":"ticketID","value":"' + $ticketToGet + '"}'
    }

    if ($ForMeOnly) {
        $istr = Get-ATEngineerToReport
        if ($istr) {
            $filters += '{"op":"eq","Field":"resourceID","value":"' + (Get-ATEngineerToReport) + '"}'
        }
        else {
            # Write-Host "No locally saved EngineerToReport details that matched existing AutoTaskResources, please run Set-ATEngineerToReport"
            throw   "No locally saved EngineerToReport details that matched existing AutoTaskResources, please run Set-ATEngineerToReport"
        }
    }
    elseif ($null -ne $ForResouerceID) {
        # $searchby = 
        $filters += '{"op":"eq","Field":"resourceID","value":"' + $ForResourceID + '"}'
    }
#    elseif ($null -ne $ForCustomerID) {
#         # $searchby = 
#         $filters += '{"op":"eq","Field":"CompanyID","value":"' + $ForCustomerID + '"}'
#     }




    # if ($ForCalendar -or $ForCalendarLiteView) {
    if ($ForCalendar ) {
        $filters += '{"op":"Exist","Field":"TicketID"}'
    }

    [string]$searchby2 = $null
    if ($filters.Count -gt 1   ) {
        $searchby2 = '{"op":"and","items":[' + ($filters -join ",") + ']}'
    }
    else {
        $searchby2 = $filters
    }


    if ($ReturnAllFields  ) {
        $timeentries = Invoke-ATQuery -entityName 'v1.0/TimeEntries' -SearchFurtherBy $searchby2 -SearchFirstBy Nothing
    }
    else {
        $timeentries = Invoke-ATQuery -entityName 'v1.0/TimeEntries' -SearchFurtherBy $searchby2 -SearchFirstBy Nothing -includeFields $includefields 

    }
     
    if (!$timeentries) {
        Write-Host "No time entries found for the given criteria." -ForegroundColor Yellow
        return
    }
    else {
        write-verbose "Read-utotaskTimeEntries: polled Autotask for the time entries -now abput to process them"
    }
    
    
    if ($ReturnRaw) {
        return $timeentries
    }
    
    
    write-host "Found $($timeentries.Count) time entries$(if ($MonthStartSTr) { " from on or after $MonthStartSTr" })" -ForegroundColor Green

    $timeentries | foreach-object {
        $_ | Add-Member -NotePropertyName "startDateTimeLocal" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "endDateTimeLocal" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "hoursNonBillable" -NotePropertyValue 0 -Force
        $_ | Add-Member -NotePropertyName "hoursBillable" -NotePropertyValue 0 -Force
        $_ | Add-Member -NotePropertyName "hoursLeave" -NotePropertyValue 0 -Force
        $_ | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force


    }

    # $timeentries | Add-Member -NotePropertyName "startDateTimeLocal" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "endDateTimeLocal" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "hoursNonBillable" -NotePropertyValue 0 -Force
    # $timeentries | Add-Member -NotePropertyName "hoursBillable" -NotePropertyValue 0 -Force
    # $timeentries | Add-Member -NotePropertyName "hoursLeave" -NotePropertyValue 0 -Force
    # $timeentries | Add-Member -NotePropertyName "IsBillableClient" -NotePropertyValue $false -Force
    # $timeentries | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force



    #Now that we have the time entries, we want to add some calculated fields to make it easier to filter and group by things such as billable vs non billable, after hours vs normal hours, etc.   
    #not needed for Calendar view, as we are not showing billable vs non billable hours, and we want to show all entries regardless of billing code, so we don't want to add the complexity of trying to determine which entries are billable vs non billable based on billing code, as this is only really relevant for productivity calculations which we are not doing in calendar view
    #TODO: check what how not getting this impacts other extracts
    #if ( !$ForCalendarLiteView) {
    # if ($IncludeBillingDetails) {
    $billingCodes = Get-ATBillingCodes 
    #  }

    write-verbose "Read-ATTimeEntries:  polling autotask for ALL Engineer details"
    # if ($includeEngineerDetails) {       
    $Engineers = Get-ATEngineers
    $timeentries | group-object ResourceID | ForEach-Object {
        $codeID = $_.Name
        $codename = ($Engineers | where-object ID -eq $codeID | select-object -first 1).Fullname
        $_.Group | Add-Member -NotePropertyName Engineer -NotePropertyValue $codename -Force
    }
    # }
    #}


    write-verbose "Get-ATTimeEntries: polling Autotask for related ticket and company info"
    #if ($includeTicketDetails) {

    $timeentries | foreach-object {
        $_ | Add-Member -NotePropertyName "TicketNumber" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "Title" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "TicketBillingCodeID" -NotePropertyValue $null -Force  
        $_ | Add-Member -NotePropertyName "CompanyID" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "CompanyName" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "CompanyClassification" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "BillingCode" -NotePropertyValue $null -Force
        $_ | Add-Member -NotePropertyName "IsBillableClient" -NotePropertyValue $false -Force
        $_  | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force

    }

    # $timeentries | Add-Member -NotePropertyName "TicketNumber" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "Title" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "TicketBillingCodeID" -NotePropertyValue $null -Force  
    # $timeentries | Add-Member -NotePropertyName "CompanyID" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "CompanyName" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "CompanyClassification" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "BillingCode" -NotePropertyValue $null -Force
    # $timeentries | Add-Member -NotePropertyName "IsBillableClient" -NotePropertyValue $false -Force
    # $ticketGroup  | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force


    $ticketIDSToSearch = $timeentries | Where-Object ticketID -gt 0 | Select-Object -ExpandProperty ticketID -Unique |  Sort-Object ticketID
    write-verbose "Get-ATTimeEntries: Polling tickets for detail" #-ForegroundColor Green
    $TicketsRetreived = Get-ATTickets -ids $ticketIDSToSearch -ForCalendar -GetCompanyNames -DontexpandticketInformation
    write-host "Finished retreiving $($TicketsRetreived.Count) tickets, now adding details to time entries" -ForegroundColor cyan
    # return

    $timeentries | Group-Object ticketID | ForEach-Object {
        $ticketID = $_.Name
        $ticketGroup = $_.Group
        if ($ticketID) {
            #  $ticket = Get-ATTickets -ids $ticketID -ForCalendar      -DontexpandticketInformation
            $ticket = $TicketsRetreived | Where-Object ID -eq $ticketID | Select-Object -First 1
            if ($ticket) {
                # if ($IncludeBillingDetails) {
                $entryTicketBillingCode = ($billingCodes | Where-Object id -eq $ticket.BillingCodeID | Select-Object -First 1).billingCode
                # } 

                $entrTicketNonBillState = (($ticket.BillingCodeID -in $ATnonBillableCodes) -or ($ticket.CompanyIsInternal))

                foreach ($entry in $ticketGroup) {
                    $entry.ticketNumber = $ticket.ticketNumber
                    $entry.Title = $ticket.title
                    $entry.CompanyID = $ticket.companyID
                    $entry.CompanyName = $ticket.CompanyName
                    $entry.CompanyClassification = $ticket.CompanyClassification
                    $entry.TicketBillingCodeID = $ticket.BillingCodeID
                    $entry.BillingCode = $entryTicketBillingCode
                    $entry.isnonBillable = $entrTicketNonBillState
                    $entry.isBillableClient = -not $ticket.CompanyIsInternal

                    # if ($IncludeBillingDetails) {
                    #                       $entry.BillingCode = $entryBillingCode

                    # }   
                }

                # $ticketGroup | Add-Member -NotePropertyName "TicketNumber" -NotePropertyValue $ticket.ticketNumber -Force
                # $ticketGroup | Add-Member -NotePropertyName "Title" -NotePropertyValue $ticket.title -Force
                # $ticketGroup | Add-Member -NotePropertyName "CompanyID" -NotePropertyValue $ticket.companyID -Force
                # $ticketGroup | Add-Member -NotePropertyName "TicketBillingCodeID" -NotePropertyValue $ticket.BillingCodeID -Force
                # $ticketGroup | Add-Member -NotePropertyName "CompanyName" -NotePropertyValue "$($ticket.CompanyName)" -Force
                # $ticketGroup | Add-Member -NotePropertyName "CompanyClassification" -NotePropertyValue "$($ticket.CompanyClassification)" -Force


                #if the ticket is non billable then ensure all sub entries are ALSO non billable, as sometimes the time entry may be marked as billable but the ticket is non billable, so we want to make sure all entries for a non billable ticket are marked as non billable to avoid confusion when filtering by billable vs non billable entries
                # if (($ticket.BillingCodeID -in $ATnonBillableCodes) -or ($ticket.CompanyIsInternal)) {
                #     #  $ticketGroup | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force
                #     $ticketGroup | foreach-object {
                #         $_.isNonBillable = $true
                #     }
                # }

                if ($ticket.CompanyIsInternal) {
                    $ticketGroup  | ForEach-Object {  
                        # $ticketGroup | Add-Member -NotePropertyName "IsBillableClient" -NotePropertyValue $false -Force
                        # $ticketGroup  | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force
                        #  $_.isbillableClient = $false  
                        #  $_.isNonBillable = $true
                        if ($_.TicketID -gt 0) {
                            $_.TicketBillingCodeID = $ATnonBillableCodes[0]
                            If ($_.Title -like "RMM*") {
                                $_.BillingCode = "RMM-Internal"
                                $_.BillingCodeID = $RMMCode
                            }
                            else {
                                $_.BillingCode = "Internal Ticket - Non Billable"         
                                $_.BillingCodeID = $ATnonBillableCodes[0]
                            }
                        }
                    }
                }
                # else {
                # $ticketGroup | foreach-object {
                # write-host "Ticket $($ticket.ticketNumber) is billable for client, but has a billing code that is in the non billable list, so marking it as non billable to avoid confusion when filtering by billable vs non billable entries" -ForegroundColor Yellow    
                #   $_.IsBillableClient = $true
                              
                # if (($ticket.BillingCodeID -in $ATnonBillableCodes) -or ($ticket.CompanyIsInternal)) {
                #      $_.isnonBillable = $true
                #  }
                # else {
                #      $_.isNonBillable = $false
                # } 
                # }
                # } 
            }
 

        }
        else {
            # this is not a ticket : it must be an internal time entry that is not associated with a ticket, so we will mark it as non billable and with company name as internal
            foreach ($entry in $ticketGroup) {
                #$entry.CompanyName = "Internal"
                #$entry.CompanyClassification = "Internal"
                $entry.IsBillableClient = $false
                $entry.isNonBillable = $true
            }
            # $ticketGroup | Add-Member -NotePropertyName "IsBillableClient" -NotePropertyValue $false -Force
            # $ticketGroup | Add-Member -NotePropertyName "isNonBillable" -NotePropertyValue $true -Force
        }
        # write-host "\" -NoNewline
    }
    # } end of if include ticket details
    write-verbose " `nGet-ATTimeEntries: Finished polling tickets for detail"    


    # if ($DisplayCompanySummary -and $includeTicketDetails ) {
    #     write-verbose "Get-ATTimeEntries: Summary of entries for each company" #-ForegroundColor Green
    #     $CompanyIDSToSearch = $timeentries | Where-Object companyID -gt 0 | Select-Object -ExpandProperty CompanyID -Unique | Sort-Object
    #     write-verbose "Found $($CompanyIDSToSearch.Count) unique company IDs" #-ForegroundColor Green
    #     $timeentries | group-object CompanyName | ForEach-Object {
    #         $companyName = $_.Name
    #         $CompanyGroup = $_.Group
    #         $hoursWorked = ($CompanyGroup | Measure-Object -Property hoursWorked -Sum).Sum
    #         if ($companyName) {
    #             write-host "Company: $($companyName) : has $($_.Count) entries, with total $hoursWorked hours" -ForegroundColor Green
    #         }
    #         else {

    #             write-host "Internal efforts (not on a ticket) : $($_.Count) entries, with total $hoursWorked hours" -ForegroundColor Green
                    

    #         }
    #     }
    # }

    #         if ($DisplayCompanySummary -and $includeTicketDetails ) {
    #         $timeentries | Where-Object IsBillableClient -eq 0 | sort-object KissWorkType | group-object KissWorkType | ForEach-Object {
    #         $workType = $_.Name
    #         $workTypeGroup = $_.Group
    #         $hoursWorked = ($workTypeGroup | Measure-Object -Property hoursWorked -Sum).Sum
    #         if ($workType) {
    #             write-host "$($workType) : has $($_.Count) entries, with total $hoursWorked hours" -ForegroundColor Green
    #         }
    #         else
    #         {
    #             write-host "Unclassified work : $($_.Count) entries, with total $hoursWorked hours" -ForegroundColor Green
    #         }
    #     }

    #     write-verbose "Get-ATTimeEntries: Summary of entries for each company" #-ForegroundColor Green
    #     $CompanyIDSToSearch = $timeentries | Where-Object companyID -gt 0 | Select-Object -ExpandProperty CompanyID -Unique | Sort-Object
    #     write-verbose "Found $($CompanyIDSToSearch.Count) unique company IDs" #-ForegroundColor Green
        

    #     $timeentries | where-object HrsClient -gt 0 | sort-object CompanyName | group-object CompanyName | ForEach-Object {
    #         $companyName = $_.Name
    #         $CompanyGroup = $_.Group
    #         $hoursWorked = ($CompanyGroup | Measure-Object -Property hoursWorked -Sum).Sum
    #         if ($companyName) {
    #             write-host "Company: $($companyName) : has $($_.Count) entries, with total $hoursWorked hours" -ForegroundColor cyan
    #         }
    #         else {

    #             write-host "Unkonwn Error - CompanyName not displayed : $($_.Count) entries, with total $hoursWorked hours" -ForegroundColor cyan
                    

    #         }
    #     }
    # }



    #process billing codes and mark non billable entries, and to add billing code names for easier filtering and grouping by billing code, such as for after hours vs normal hours, or for internal projects vs external projects, etc.
    write-verbose "Adjusting Billing codes and optionally adding text description to them"
   



    $AllleaveCodes = $LeaveCodes + $SickCodes #
    $timeentries | group-object BillingCodeID | ForEach-Object {
        $codeID = $_.Name
        $codeGroup = $_.Group
        if ($codeID -in $ATnonBillableCodes) {
            $codeGroup | Add-Member -NotePropertyName isNonBillable -NotePropertyValue $true -Force 
            # $codeGroup | ForEach-Object {
            #     if ( $codeGroup) {
            #      }                
            # }
        }
        # if ($IncludeBillingDetails) {
        if ($codeGroup) {
            $codename = ($billingCodes | where-object ID -eq $codeID | select-object -first 1).name
            $codeGroup | Add-Member -NotePropertyName BillingCode -NotePropertyValue $codename -Force 
        }   
        # $codename = ($billingCodes | where-object ID -eq $codeID | select-object -first 1).name
        # $_.Group | Add-Member -NotePropertyName BillingCode -NotePropertyValue $codename -Force
        # }

        if ($codeID -in $afterHrsBillingCodes) {
            if ($codeGroup) {
                $codeGroup | Add-Member -NotePropertyName isAfterHours -NotePropertyValue $true -Force 
            }
            # $codeGroup | ForEach-Object {
            #     if ( $codeGroup) {
            #      #  
            #        $_ | Add-Member -NotePropertyName isAfterHours -NotePropertyValue $true -Force 
            #     }                
            # }
        }

        if ($codeID -in $AllleaveCodes) {
            #the entry is Leave or Sick leave - then we want to mark the entry as leave and also add the hours worked to a separate field called hoursLeave, so that we can easily filter and group by leave entries, and also exclude them from productivity calculations if needed, as these are not really non billable work but rather time off that should be treated separately in reporting and analysis
            if ( $codeGroup) { 
                $codeGroup | ForEach-Object {
                    $_ | Add-Member -NotePropertyName hoursLeave -NotePropertyValue $_.HoursWorked -Force 
                }                
                $codeGroup | Add-Member -NotePropertyName hoursWorked -NotePropertyValue 0 -Force 
                $codeGroup | Add-Member -NotePropertyName hoursnonBillable -NotePropertyValue 0 -Force 
                # $codeGroup| Add-Member -NotePropertyName CompanyName -NotePropertyValue "Leave" -Force 
            }
        }
    }



    write-verbose "Add information about roles"

    #process roleID to add role names for easier filtering and grouping by role, such as for after hours vs normal hours, or for internal projects vs external projects, etc.
    #  if ($IncludeBillingDetails -and (!$ForCalendarLiteView -and !$ForCalendar)) {       
    if ( !$ForCalendar) {       
        $roles = Get-ATRoles
        $timeentries | group-object roleID | ForEach-Object {
            $roleID = $_.Name
            $codename = ($roles | where-object ID -eq $roleID | select-object -first 1).name
            #$_.Group | Add-Member -NotePropertyName Role -NotePropertyValue $codename -Force
            $_.Group | Add-Member -NotePropertyName Role -NotePropertyValue $codename -Force
        }
    }

    write-verbose "creating Local DateTimes and updating HoursBillable/hoursnonBillable"
    foreach ($entry in $timeentries) {
        if ($null -ne $entry.startDateTime) {
            $entry.startDateTimeLocal = [datetime]$entry.startDateTime.ToLocalTime()
        }
        if ($null -ne $entry.endDateTime) {
            $entry.endDateTimeLocal = [datetime]$entry.endDateTime.ToLocalTime()
        }
        if (($entry.isNonBillable -eq $true) -or ($entry.CompanyIsInternal -eq $true)) {

            $entry.hoursNonBillable = $entry.hoursWorked
            $entry.hoursBillable = 0 
        }
        else {
            $entry.hoursNonBillable = 0
            $entry.hoursBillable = $entry.hoursWorked      
        }
    }


    

    # $timeentries
    # return
        
    write-verbose " either return now for Calendaar - or sort entries by endtime or OADate"
    #    if ($ForCalendar -or $ForCalendarLiteView) {
    if ($ForCalendar ) {
        # Write-Host "Get-ATTimeEntries: Done Processing time entries for calendar view" -ForegroundColor Green
        return $timeentries  | sort-object enddateTime 
    }
    else {
        #create a numerically sortable date field
        $timeentries | Add-Member -NotePropertyName 'OADate' -NotePropertyValue 0.0
        foreach ($i in $timeentries) {
            $i.OADate = ([datetime]$i.dateWorked).ToOADate()
        }

        #         $utcString = "3/03/2026 7:45:00 pm"

        # # Convert string → DateTime (tell PowerShell it's UTC)
        # $utcTime = [datetime]::Parse($utcString, $null, [System.Globalization.DateTimeStyles]::AssumeUniversal)
        # # Convert UTC → Local
        # $localTime = $utcTime.ToLocalTime()
        # $localTime



        write-verbose "Creating searchable date strings"
        Convert-ObjArrayDateTimesToSearchableStrings -obj $timeentries 
   

        #    write-verbose "Get-ATTimeEntries count = $($timeentries.count)"
        # Now provide calculate Columns to assist with stats
        # $timeentries | Add-Member -NotePropertyName 'OADate' -NotePropertyValue 0.0
        foreach ($i in $timeentries) {
            $i.OADate = ([datetime]$i.dateWorked).ToOADate()
            $i | Add-Member -NotePropertyName 'kissWorkType' -NotePropertyValue ""
            $i | Add-Member -NotePropertyName 'HrsClientBillableNormalHrs' -NotePropertyValue 0.0  
            $i | Add-Member -NotePropertyName 'HrsClientBillableAfterHrs' -NotePropertyValue 0.0 
            $i | Add-Member -NotePropertyName 'HrsClientNonBillableNormalHrs' -NotePropertyValue 0.0 
            $i | Add-Member -NotePropertyName 'HrsClientNonBillableAfterHrs' -NotePropertyValue 0.0 
            $i | Add-Member -NotePropertyName 'HrsClient' -NotePropertyValue 0.0 

            $i | Add-Member -NotePropertyName 'HrsLeave' -NotePropertyValue 0.0 
            $i | Add-Member -NotePropertyName 'HrsSick' -NotePropertyValue 0.0
            $i | Add-Member -NotePropertyName 'HrsTeaBreaks' -NotePropertyValue 0.0
            $i | Add-Member -NotePropertyName 'HrsTraining' -NotePropertyValue 0.0
            $i | Add-Member -NotePropertyName 'HrsInternalProd' -NotePropertyValue 0.0
            $i | Add-Member -NotePropertyName 'HrsInternalOther' -NotePropertyValue 0.0 
            $i | Add-Member -NotePropertyName 'AfterHours' -NotePropertyValue 0.0
        }
        # $timeentries | Add-Member -NotePropertyName 'kissWorkType' -NotePropertyValue ""
        #     $timeentries | Add-Member -NotePropertyName 'HrsClientBillableNormalHrs' -NotePropertyValue 0.0  
        # $timeentries | Add-Member -NotePropertyName 'HrsClientBillableAfterHrs' -NotePropertyValue 0.0 
        # $timeentries | Add-Member -NotePropertyName 'HrsClientNonBillableNormalHrs' -NotePropertyValue 0.0 
        # $timeentries | Add-Member -NotePropertyName 'HrsClientNonBillableAfterHrs' -NotePropertyValue 0.0 
        # $timeentries | Add-Member -NotePropertyName 'HrsClient' -NotePropertyValue 0.0 

        # $timeentries | Add-Member -NotePropertyName 'HrsLeave' -NotePropertyValue 0.0 
        # $timeentries | Add-Member -NotePropertyName 'HrsSick' -NotePropertyValue 0.0
        # $timeentries | Add-Member -NotePropertyName 'HrsTeaBreaks' -NotePropertyValue 0.0
        # $timeentries | Add-Member -NotePropertyName 'HrsTraining' -NotePropertyValue 0.0
        # $timeentries | Add-Member -NotePropertyName 'HrsInternalProd' -NotePropertyValue 0.0
        # $timeentries | Add-Member -NotePropertyName 'HrsInternalOther' -NotePropertyValue 0.0 
        # $timeentries | Add-Member -NotePropertyName 'AfterHours' -NotePropertyValue 0.0
    
    
  


        write-verbose " processing the detailed hrs summary"
        #---------------
        #Process the Ticket (customer related) time entries
        $subitems = $timeentries | Where-Object ticketID 
        if ($subitems) {
            
            $internalTicketItems = $subitems | Where-Object { ($_.CompanyClassification -eq $ATInternalClasificationCode) }
            if ($internalTicketItems) {
                $internalTicketItems | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Internal-Ticket" -Force
                foreach ($item in $internalTicketItems) {
                    $item.HrsInternalProd = $item.hoursWorked
                    $item.kissWorkType = "Internal-Ticket"
                    # $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Internal-Ticket" -Force

                    # $item.isNonBillable = $true
                    # $item.ticket = $item.hoursWorked
                } 
            }
            $Clientitems = $subitems | Where-Object { ($_.CompanyClassification -ne $ATInternalClasificationCode) }

            if ($Clientitems) {
                $items = $Clientitems | Where-Object { ($_.isNonBillable -eq $true) }
                if ($items) {
                    $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Client-NonBillable-NormalHrs" -Force
                    foreach ($item in $items) {
                        $item.HrsClientNonBillableNormalHrs = $item.hoursWorked
                        $item.hrsClient = $item.hoursWorked
                    } 
                }
     
                $items = $Clientitems | Where-Object { ($_.isNonBillable -ne $true) }
                if ($items) {
                    $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Client-Billable-NormalHrs" -Force
                    foreach ($item in $items) {
                        $item.HrsClientBillableNormalHrs = $item.hoursWorked
                        $item.hrsClient = $item.hoursWorked
                    }   
                }

                #identify the afterhours billable
                $items = $subitems | Where-Object { ($_.billingCodeID -in $afterHrsBillingCodes) -and ($_.isNonBillable -ne $true) }
                if ($items) {
                    $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Billable-AfterHrs" -Force
                    foreach ($item in $items) {
                        $item.HrsClientBillableAfterHrs = $item.hoursWorked
                        $item.HrsClientBillableNormalHrs = 0
                        $item.hrsClient = $item.hoursWorked
                        $item.afterhours = $item.hoursWorked
                    }

                }
                #identify the afterhours nonbillable
                $items = $subitems | Where-Object { ($_.billingCodeID -in $afterHrsBillingCodes) -and ($_.isNonBillable -eq $true) }
                if ($items) {
                    $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Non-Billable-AfterHrs" -Force
                    foreach ($item in $items) {
                        $item.HrsClientNonBillableAfterHrs = $item.hoursWorked
                        $item.HrsClientNonBillableNormalHrs = 0
                        $item.hrsClient = $item.hoursWorked
                        $item.afterhours = $item.hoursWorked
                    }

                }
            }
        }

        else { Write-Verbose "No Client ticket items found in timesheet entries" }


        write-verbose "Processing the Interal entries"
        #return $timeentries
        #------------------------------
        # now process all the Internal, leave, admin etc
        $subitems = $timeentries | Where-Object { !($_.ticketID ) }
        if ($subitems) {

            # set default for ALL internal work that it is non billable  and not personal    
            $items = $subitems | Where-Object { ($_.billingCodeID -notin $leaveCodes, $sickCodes) }
            if ($items) {
                $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Internal-Other" -Force
                foreach ($item in $items) {
                    $item.HrsInternalOther = $item.hoursWorked
                    # $item.HrsNonStatistic = 0.0
                    # $item.HrsAfterHrs = 0.0
                    # $item.kissWorkType = "Internal-NonBillable"
                } 
            }

            $items = $subitems | Where-Object { ($_.billingCodeID -in $LeaveCodes) }
            if ($items) {
                $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Leave" -Force
                foreach ($item in $items) {
                    #  $item.HrsNormal = 0.0
                    #  $item.HrsNonStatistic = $item.hoursWorked
                    $item.hrsleave = $item.hoursWorked
                    # $item.kissWorkType = "Leave"
                    $item.HrsInternalOther = 0
                }
            }

            $items = $subitems | Where-Object { ($_.billingCodeID -in $SickCodes) }
            if ($items) {
                $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Sick" -Force
                foreach ($item in $items) {
                    # $item.HrsNormal = 0.0
                    # $item.HrsNonStatistic = $item.hoursWorked
                    $item.HrsSick = $item.hoursWorked
                    # $item.kissWorkType = "Sick"
                    $item.HrsInternalOther = 0
                }
            }
            $items = $subitems | Where-Object { ($_.billingCodeID -in $TrainingCodes) }
            if ($items) {
                $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Training" -Force
                foreach ($item in $items) {
                    # $item.HrsNormal = $item.hoursWorked
                    $item.HrsTraining = $item.hoursWorked
                    $item.HrsInternalOther = 0
                }
            }
            $items = $subitems | Where-Object { ($_.billingCodeID -in $teabreakCodes) }
            if ($items) {

                $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "TeaBreaks" -Force
                foreach ($item in $items) {
                    # $item.HrsNormal = $item.hoursWorked
                    $item.HrsTeaBreaks = $item.hoursWorked
                    $item.HrsInternalOther = 0
                }
            }

            $items = $subitems | Where-Object { ($_.billingCodeID -in $ProductiveCodes) }
            if ($items) {
                $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Internal-Prod" -Force
                foreach ($item in $items) {
                    $item.HrsInternalProd = $item.hoursWorked
                    $item.HrsInternalOther = 0
                }
            }
        }
        else { Write-Verbose "No Internal items found in timesheet entries" }

        write-verbose "building the RMM time info"
        $timeEntries | Add-Member -NotePropertyName 'RMMTicket' -NotePropertyValue 0.0 -Force
        $timeEntries | Add-Member -NotePropertyName 'RMMTask' -NotePropertyValue 0.0 -Force
        $items = $timeEntries | Where-Object Title -Like "RMM*"
        foreach ($item in $items) {
            $item.RMMTicket = $item.hoursWorked
            Write-debug "Build-AutoTaskRMMTickets: found RMM ticket time entry $($item.hoursworked)  on $($RMMTickets.id) $($timeEntries.title) "
        }  
        foreach ($rmm in $ATtaskCodesRMM) {
            $RMMtasks = $timeEntries | Where-Object BillingCodeID -eq $rmm
            foreach ($task in $RMMtasks) {
                $task.RMMTask = $task.hoursWorked
            }        
        }

        #Set-ATInternalTicketTime $timeentries | Out-Null
        #Build-AutoTaskRMMTime $timeentries | Out-Null
        #write-Host "DONE polling AutoTask for TimeEntries, and formating the results" -foregroundcolor green




        $timeentries

        # if ($DisplayCompanySummary -and $includeTicketDetails ) {
        if ($DisplayCompanySummary  ) {
            $timeentries | Where-Object IsBillableClient -ne $true | sort-object KissWorkType | group-object KissWorkType | ForEach-Object {
                $workType = $_.Name
                #write-Host "WorkType: $workType" -ForegroundColor Green
                $workTypeGroup = $_.Group
                $hoursWorked = ($workTypeGroup | Measure-Object -Property hoursWorked -Sum).Sum
                if ($workType -like "Leave") {
                    $leave = ($workTypeGroup | Measure-Object -Property hoursLeave -Sum).SUm
                    write-Host "Leave (including Sick Leave) =  $Leave hours"
                }
                elseif ($workType -notlike "Leave") {
                    write-host "$($workType.PadRight(30)) : hours = $hoursWorked " -ForegroundColor Green
                }
                else {
                    write-host "Unclassified work : hours =   $hoursWorked" -ForegroundColor Green
                }
            }

            write-verbose "Get-ATTimeEntries: Summary of entries for each company" #-ForegroundColor Green
            #$CompanyIDSToSearch = $timeentries | Where-Object companyID -gt 0 | Select-Object -ExpandProperty CompanyID -Unique | Sort-Object
            #write-verbose "Found $($CompanyIDSToSearch.Count) unique company IDs" #-ForegroundColor Green
        

            $timeentries | where-object isBillableClient -eq $true | sort-object CompanyName | group-object CompanyName | ForEach-Object {
                $companyName = $_.Name
                $CompanyGroup = $_.Group
                $hoursWorked = ($CompanyGroup | Measure-Object -Property hoursWorked -Sum).Sum
                if ($companyName) {
                    write-host "Company: $($companyName.PadRight(30)) : $($_.Count.tostring().PadLeft(3)) entries, hours = $hoursWorked" -ForegroundColor cyan
                }
                else {

                    write-host "Unkonwn Error - CompanyName not displayed : $($_.Count.tostring().PadLeft(3)) with total $hoursWorked hours" -ForegroundColor cyan
                }
            }
        }

    } 

}


function Get-ATWeeklySummary {
    <#
    .SYNOPSIS
    Prints and returns a weekly hours summary grouped by engineer.

    .DESCRIPTION
    Retrieves time entries for the last N weeks (starting from the previous Sunday) and
    produces a per-engineer summary of billable, non-billable client, and internal productive hours.

    Results are both printed to the host in green and returned as a PSCustomObject array
    with the properties: Engineer, hoursBillable, hoursNonBillable, hoursInternal, weeks,
    startDate, and endDate.

    .PARAMETER lastXWeeks
    How many weeks back to summarise. Each week starts on a Sunday.
    Default is 1 (the current week from the last Sunday to now).

    .PARAMETER AllEngineers
    Switch. When set, retrieves time entries for all engineers rather than only the
    currently saved engineer (ForMeOnly). Useful for team-level reporting.

    .EXAMPLE
    Get-ATWeeklySummary
    # Summary for the current week for the saved engineer

    .EXAMPLE
    Get-ATWeeklySummary -lastXWeeks 2 -AllEngineers
    # Two-week summary for every engineer

    .NOTES
    Calls Get-ATTimeEntries internally. Run time depends on the number of entries
    in the requested period — typically a few minutes for a full team over 2+ weeks.
    The Engineer grouping uses the 'Engineer' property added by Get-ATTimeEntries.
    #>
    [CmdletBinding()]
    param (
       
        [Parameter(Position=0)]
        [int]$lastXWeeks = 1,
        [switch]$AllEngineers = $false
    )
    
    if ($lastXWeeks -gt 0){
        $beforedate  = (Get-ATWeekStart -LastXWeeks 0)
    }
    else
    {
    $beforedate = get-date -Hour 00 -Minute 0 -Second 0
    $beforedate = $beforedate.AddDays(1)

    }

    # $FirstSunday = Get-ATWeekStart -LastXWeeks $lastXWeeks
    if ($AllEngineers) {
        $timeentries = Get-ATTimeEntries -LastXWeeks $lastXWeeks -UntilDateLocal $beforedate -includeLeave
    }
    else {
        $timeentries = Get-ATTimeEntries -LastXWeeks $lastXWeeks   -ForMeOnly -IncludeSummaryNotes -UntilDateLocal $beforedate -includeLeave
    }
    write-host "`nGot $($timeentries.count) time entries, now summarising " -ForegroundColor Cyan
    $startDate = $timeentries | Measure-Object -Property dateWorked -Minimum | Select-Object -ExpandProperty Minimum
    $endDate = $timeentries | Measure-Object -Property dateWorked -Maximum | Select-Object -ExpandProperty Maximum
    #  $days = ($endDate - $startDate).Days + 1
    $summaries = @()

    $timeentries | group-object Engineer  | ForEach-Object { 
        $engineer = $_.Name
        $hoursClient = [Math]::Round(($_.Group | Measure-Object -Property HrsClient -Sum).Sum, 2)
        $hoursBillable = [Math]::Round(($_.Group  | Measure-Object -Property hoursBillable -Sum).Sum, 2)  
        $hoursAfterHours = [Math]::Round(($_.Group  | Measure-Object -Property AfterHours -Sum).Sum, 2) 
        $hoursAdmin = [Math]::Round(($_.Group | Where-object { $null -eq $_.TicketID } | Measure-Object -Property hoursworked -Sum).Sum, 2) 
        $hoursTimesheeted = [Math]::Round(($_.Group | Measure-Object -Property hoursWorked -Sum).Sum, 2)
        $hoursLeave = [Math]::Round(($_.Group | Measure-Object -Property hoursLeave -Sum).Sum, 2)
        $hoursNonBillable = $hoursClient - $hoursBillable
        $hoursInternal = [Math]::Round(($_.Group  | Measure-Object -Property HrsInternalProd -Sum).Sum, 2)
        $percentBillable = [Math]::Round(($hoursClient *100)/$hoursTimesheeted,0)
        $percentproductive = [Math]::Round((($hoursClient + $hoursInternal)*100)/$hoursTimesheeted,0)

        #$hoursClientBillableNormal = [Math]::Round(($_.Group | Measure-Object -Property HrsClientBillableNormalHrs -Sum).Sum, 2)
        #$HrsClientBillableAfterHrs  = [Math]::Round(($_.Group | Measure-Object -Property HrsClientBillableAfterHrs -Sum).Sum, 2)
         Write-Host "Summary for Engineer: $engineer for the last $lastXWeeks weeks  (from $($startDate.ToShortDateString()) to $($endDate.ToShortDateString())):" -ForegroundColor Green
        # Write-Host "Summary for Engineer: $engineer for the last $lastXWeeks weeks (from $startDate to $endDate):" -ForegroundColor Green
        if ($hoursAfterHours -gt 0) {
            Write-Host "-Billable Hours: $hoursBIllable *includes* AfterHours: $hoursAfterHours" -ForegroundColor Green
        }
        else {
            Write-Host "-Billable Hours           : $hoursBIllable" -ForegroundColor Green
        }
            Write-Host "-Non-Billable Client Hours: $hoursNonBillable" -ForegroundColor Green
            write-host "-Internal Tickets         : $hoursInternal and  Admin: $hoursAdmin"  -ForegroundColor Green
        if ($hoursLeave) {
            write-host "-Total hours timesheeted  : $hourstimesheeted  (plus Leave Hours: $hoursLeave )"  -ForegroundColor Green
        }
        else {
            write-host "-Total hours timesheeted  : $hourstimesheeted"  -ForegroundColor Green
        }
            write-host "- % Billable              : $percentBillable" -ForegroundColor cyan
            write-host "- % productive            : $percentproductive" -ForegroundColor cyan


        $outp = [PSCustomObject]@{
            Engineer    = $engineer
            Billable    = $hoursBillable
            NonBillable = $hoursNonBillable
            Internal    = $hoursInternal + $hoursAdmin
            TimeSheeted = $hourstimesheeted
            weeks       = $lastXWeeks
            Leave       = $hoursLeave
            startDate   = $startDate
            endDate     = $endDate
            percentBillable = $percentBillable
            percentproductive = $percentproductive
        }
        $summaries += $outp
    }
    $summaries

}

function Get-ATEffortSummary {
    <#
    .SYNOPSIS
    Produces a detailed effort and time-spent report grouped by company and ticket,
    including internal work and entries that have no start/end time recorded.

    .DESCRIPTION
    Retrieves time entries for the specified period and builds a two-level summary:

      1. A per-ticket / per-admin row showing: company, ticket number, title, engineer,
         date worked, hours, whether billable, and whether the entry is internal.
         Entries that have no startDateTime / endDateTime (i.e. duration-only entries
         recorded without a clock start/stop) are included and flagged with
         HasTimestamp = $false so they are visible but distinguishable.

      2. A per-company rollup showing total billable, non-billable, internal, and
         admin (no-ticket) hours — printed to the console as a readable summary and
         also returned as objects.

    Internal entries (company classification = $CONSTInternalClasificationCode, or
    entries with no ticket that use internal billing codes) are included in full and
    marked with IsInternal = $true so they can be filtered or highlighted downstream.

    .PARAMETER LastXWeeks
    How many weeks back to include, counting from the previous Sunday. Default is 1.

    .PARAMETER LastxMonths
    Alternative to LastXWeeks — how many calendar months back to start from.
    Ignored when LastXWeeks > 0.

    .PARAMETER FromDateLocal
    Explicit local start date. Overrides LastXWeeks and LastxMonths.

    .PARAMETER ForMeOnly
    Switch. When set, only retrieves entries for the engineer saved via Set-ATEngineerToReport.

    .PARAMETER AllEngineers
    Switch. When set, retrieves entries for all engineers. Mutually exclusive with ForMeOnly.
    Default behaviour (neither switch) also returns all engineers.

    .PARAMETER IncludeLeaveAndBreaks
    Switch. When set, leave, sick, and tea-break entries are included in the detail rows.
    By default these are excluded from the per-ticket detail but are still counted in the
    per-company rollup totals so the hours always reconcile.

    .PARAMETER PassThruEntries
    Switch. When set, the raw enriched time-entry array is also returned alongside the
    summary objects, wrapped in a PSCustomObject with properties Detail and Summary.
    Without this switch only the summary rows are returned.

    .EXAMPLE
    # Current week, my entries only — print summary to console and return objects
    Get-ATEffortSummary -ForMeOnly

    .EXAMPLE
    # Last 2 weeks, all engineers
    Get-ATEffortSummary -LastXWeeks 2 -AllEngineers

    .EXAMPLE
    # Last month, return raw entries too
    $report = Get-ATEffortSummary -LastxMonths 1 -PassThruEntries
    $report.Summary | Export-Csv .\EffortSummary.csv -NoTypeInformation
    $report.Detail  | Export-Csv .\EffortDetail.csv  -NoTypeInformation

    .EXAMPLE
    # From a specific date forward, include leave entries
    Get-ATEffortSummary -FromDateLocal ([datetime]'2026-03-01') -IncludeLeaveAndBreaks

    .NOTES
    Calls Get-ATTimeEntries internally with -IncludeSummaryNotes, so it performs
    several API calls and may take a few minutes on large date ranges.

    Entries with no startDateTime/endDateTime are those where the engineer entered hours as
    a flat duration rather than clocking in and out. They are fully included — only the
    HasTimestamp flag distinguishes them.

    The IsInternal flag is $true when:
      - The ticket belongs to a company whose classification matches $CONSTInternalClasificationCode, OR
      - The entry has no ticket AND its billing code is in a non-client category
        (leave, sick, tea breaks, training, internal-productive codes).
    #>
    [CmdletBinding()]
    param (
        [int]$LastXWeeks = 0,
        [int]$LastxMonths = 1,
        [Nullable[DateTime]]$FromDateLocal = $null,

        [switch]$ForMeOnly = $false,
        [switch]$AllEngineers = $false,
        [switch]$IncludeLeaveAndBreaks = $false,
        [switch]$PassThruEntries = $false,

        # Billing code sets — keep defaults aligned with Get-ATTimeEntries
        [int[]]$LeaveCodes = @(91206, 29718729),
        [int[]]$SickCodes = @(91207),
        [int[]]$TeaBreakCodes = @(91209),
        [int[]]$TrainingCodes = @(29683344),
        [int[]]$ProductiveCodes = @(29711172, 29712660, 29713657, 29737360, 29718730)
    )

    # ── 1. Fetch enriched time entries ────────────────────────────────────────
    $fetchParams = @{
        # IncludeBillingDetails  = $true
        # includeTicketDetails   = $true
        # includeEngineerDetails = $true
        IncludeSummaryNotes = $true
    }

    if ($null -ne $FromDateLocal) {
        $fetchParams['FromDateLocal'] = $FromDateLocal
    }
    elseif ($LastXWeeks -gt 0) {
        $fetchParams['LastXWeeks'] = $LastXWeeks
    }
    else {
        $fetchParams['LastxMonths'] = $LastxMonths
    }

    if ($ForMeOnly) {
        $fetchParams['ForMeOnly'] = $true
    }

    Write-Host "Get-ATEffortSummary: fetching time entries..." -ForegroundColor Cyan
    $entries = Get-ATTimeEntries @fetchParams

    if (-not $entries) {
        Write-Host "Get-ATEffortSummary: no time entries found for the given criteria." -ForegroundColor Yellow
        return
    }
    Write-Host "Get-ATEffortSummary: processing $($entries.Count) entries." -ForegroundColor Cyan

    # ── 2. Internal billing-code sets (no-ticket entries) ────────────────────
    $internalNonTicketCodes = $LeaveCodes + $SickCodes + $TeaBreakCodes + $TrainingCodes + $ProductiveCodes

    # ── 3. Build detail rows ──────────────────────────────────────────────────
    $detail = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($e in $entries) {

        # Determine whether this entry has a clock-in/clock-out timestamp
        $hasTimestamp = ($null -ne $e.startDateTime) -and ($null -ne $e.endDateTime)

        # Determine whether internal
        $isInternal = $false
        if ($e.CompanyIsInternal -eq $true) {
            $isInternal = $true
        }
        elseif (-not $e.ticketID -and ($e.billingCodeID -in $internalNonTicketCodes)) {
            $isInternal = $true
        }

        # Optionally skip leave/break entries from detail (still counted in rollup)
        $isLeaveOrBreak = (-not $e.ticketID) -and ($e.billingCodeID -in ($LeaveCodes + $SickCodes + $TeaBreakCodes))
        if ($isLeaveOrBreak -and -not $IncludeLeaveAndBreaks) { continue }

        # Resolve start/end for display — use dateWorked when no timestamp exists
        $startDisplay = if ($hasTimestamp -and $e.startDateTimeLocal) {
            $e.startDateTimeLocal.ToString('yyyy-MM-dd HH:mm')
        }
        elseif ($e.dateWorked) {
            # dateWorked may be a string after Convert-ObjArrayDateTimesToSearchableStrings
            "$($e.dateWorked -replace 'T.*','')"
        }
        else { '' }

        $endDisplay = if ($hasTimestamp -and $e.endDateTimeLocal) {
            $e.endDateTimeLocal.ToString('HH:mm')
        }
        else { '' }

        $row = [PSCustomObject]@{
            Engineer       = if ($e.Engineer) { $e.Engineer }      else { "ID:$($e.resourceID)" }
            DateWorked     = $e.dateWorked -replace 'T.*', ''
            StartTime      = $startDisplay
            EndTime        = $endDisplay
            HasTimestamp   = $hasTimestamp
            Company        = if ($e.CompanyName) { $e.CompanyName }   else { '' }
            TicketNumber   = if ($e.TicketNumber) { $e.TicketNumber }  else { '' }
            Title          = if ($e.Title) { $e.Title }         else { '' }
            SummaryNotes   = if ($e.summaryNotes) { $e.summaryNotes }  else { '' }
            BillingCode    = if ($e.BillingCode) { $e.BillingCode }   else { '' }
            HoursWorked    = [Math]::Round($e.hoursWorked, 2)
            AfterHours     = $e.AfterHours
            IsBillable     = -not ($e.isNonBillable -eq $true)
            IsInternal     = $isInternal
            IsLeaveOrBreak = $isLeaveOrBreak
            WorkType       = if ($e.kissWorkType) { $e.kissWorkType }  else { 'Unknown' }
            TimeEntryID    = $e.id
        }
        $detail.Add($row)
    }

    # ── 4. Per-company rollup (uses ALL entries, not the filtered detail) ─────
    $summary = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Group by company name — entries with no ticket have null/empty CompanyName
    $grouped = $entries | Group-Object { if ($_.CompanyName) { $_.CompanyName } else { '(Internal / No Ticket)' } }

    foreach ($grp in ($grouped | Sort-Object Name)) {
        $companyName = $grp.Name
        $grpEntries = $grp.Group

        $totalHours = [Math]::Round(($grpEntries | Measure-Object hoursWorked -Sum).Sum, 2)
        $billable = [Math]::Round(($grpEntries | Where-Object { $_.isNonBillable -ne $true -and $_.IsBillableClient -eq $true } | Measure-Object hoursWorked -Sum).Sum, 2)
        $nonBillable = [Math]::Round(($grpEntries | Where-Object { $_.isNonBillable -eq $true -and $_.IsBillableClient -eq $true } | Measure-Object hoursWorked -Sum).Sum, 2)
        $internal = [Math]::Round(($grpEntries | Where-Object { $_.CompanyIsInternal -eq $true } | Measure-Object hoursWorked -Sum).Sum, 2)
        $noTicket = [Math]::Round(($grpEntries | Where-Object { -not $_.ticketID }              | Measure-Object hoursWorked -Sum).Sum, 2)
        $noTimestamp = [Math]::Round(($grpEntries | Where-Object { $null -eq $_.startDateTime }    | Measure-Object hoursWorked -Sum).Sum, 2)
        $ticketCount = ($grpEntries | Where-Object ticketID | Select-Object -ExpandProperty ticketID -Unique).Count
        $isInternalCo = ($grpEntries | Where-Object { $_.CompanyIsInternal -eq $true }).Count -gt 0

        # Console output
        $tag = if ($isInternalCo) { ' [INTERNAL]' } else { '' }
        Write-Host ""
        Write-Host "  $companyName$tag" -ForegroundColor $(if ($isInternalCo) { 'Yellow' } else { 'Cyan' })
        Write-Host "    Tickets: $ticketCount   Total hrs: $totalHours" -ForegroundColor White
        if ($billable -gt 0) { Write-Host "    Billable:     $billable hrs"     -ForegroundColor Green }
        if ($nonBillable -gt 0) { Write-Host "    Non-Billable: $nonBillable hrs"  -ForegroundColor Gray }
        if ($internal -gt 0) { Write-Host "    Internal:     $internal hrs"     -ForegroundColor Yellow }
        if ($noTicket -gt 0) { Write-Host "    Admin/No Ticket: $noTicket hrs"  -ForegroundColor Gray }
        if ($noTimestamp -gt 0) { Write-Host "    No timestamp: $noTimestamp hrs (duration-only entries)" -ForegroundColor DarkYellow }

        $summary.Add([PSCustomObject]@{
                Company          = $companyName
                IsInternal       = $isInternalCo
                TicketCount      = $ticketCount
                TotalHours       = $totalHours
                BillableHours    = $billable
                NonBillableHours = $nonBillable
                InternalHours    = $internal
                AdminNoTicket    = $noTicket
                DurationOnlyHrs  = $noTimestamp
            })
    }

    # ── 5. Grand totals ───────────────────────────────────────────────────────
    $grandTotal = [Math]::Round(($entries    | Measure-Object hoursWorked -Sum).Sum, 2)
    $grandBill = [Math]::Round(($summary    | Measure-Object BillableHours -Sum).Sum, 2)
    $grandNonBill = [Math]::Round(($summary    | Measure-Object NonBillableHours -Sum).Sum, 2)
    $grandInternal = [Math]::Round(($summary    | Measure-Object InternalHours -Sum).Sum, 2)
    $grandNoStamp = [Math]::Round(($entries    | Where-Object { $null -eq $_.startDateTime } | Measure-Object hoursWorked -Sum).Sum, 2)

    Write-Host ""
    Write-Host "─────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  GRAND TOTAL" -ForegroundColor White
    Write-Host "    All hours timesheeted : $grandTotal"    -ForegroundColor White
    Write-Host "    Billable              : $grandBill"     -ForegroundColor Green
    Write-Host "    Non-Billable (client) : $grandNonBill"  -ForegroundColor Gray
    Write-Host "    Internal              : $grandInternal" -ForegroundColor Yellow
    if ($grandNoStamp -gt 0) {
        Write-Host "    Duration-only entries : $grandNoStamp hrs (no clock start/end)" -ForegroundColor DarkYellow
    }
    Write-Host "─────────────────────────────────────────" -ForegroundColor DarkGray

    # ── 6. Return ─────────────────────────────────────────────────────────────
    if ($PassThruEntries) {
        return [PSCustomObject]@{
            Summary = $summary.ToArray()
            Detail  = $detail.ToArray()
            Entries = $entries
        }
    }
    else {
        # Return summary + detail as a flat array sorted by company then date
        # Callers can split them with Where-Object { $_.PSObject.Properties.Name -contains 'TicketNumber' }
        return [PSCustomObject]@{
            Summary = $summary.ToArray()
            Detail  = $detail.ToArray()
        }
    }
}

function Get-ATTicketReport {
    <#
    .SYNOPSIS
    Produces a comprehensive report for one or more tickets, including the ticket header,
    description, all internal and external notes, and every time entry with its summary notes.

    .DESCRIPTION
    For each requested ticket the function fetches:

      1. Ticket header — title, description, status, queue, assigned engineer, company,
         created/completed dates, and whether the ticket is billable.

      2. All ticket notes — both Internal (noteType 1) and External / client-visible (noteType 2)
         notes, sorted by creation date.  Each note shows the author, date, and full text.

      3. All time entries — every time entry posted against the ticket, including the engineer
         name, date worked, start/end times (or a flag when none were recorded), hours, billing
         code, and the engineer's summary notes.

    Tickets can be identified by their numeric Autotask IDs or by their ticket numbers
    (e.g. 'T20260101.0001').  Both parameters accept arrays so multiple tickets can be
    reported in a single call.

    The function prints a formatted console report and returns a structured object whose
    properties (Header, Notes, TimeEntries, Summary) can be piped to Export-Csv or used
    in further processing.

    .PARAMETER ids
    One or more numeric Autotask ticket IDs.

    .PARAMETER TicketNumbers
    One or more Autotask ticket numbers (e.g. 'T20260101.0001', 'T20260315.0042').

    .PARAMETER IncludeExternalNotes
    Switch. When set, client-visible (external) notes are included alongside internal notes.
    Default is internal notes only.

    .PARAMETER ExcludeInternalNotes
    Switch. When set, internal (staff-only) notes are excluded.
    Default is $false — internal notes are always included unless this is explicitly set to $false.

    .PARAMETER SuppressConsoleOutput
    Switch. When set, the formatted console report is suppressed and only the return object
    is produced. Useful when calling from scripts that process the data programmatically.

    .PARAMETER OutputHtml
    Switch. Generates a self-contained HTML report in addition to (or instead of) console
    output. Light mode (print-friendly, minimal ink) is used by default; add -DarkMode for
    the dark-themed version. The saved path is written to the console.

    .PARAMETER HtmlPath
    Full path for the HTML file. Defaults to a timestamped file in $env:TEMP.

    .PARAMETER OpenInBrowser
    Switch. Generates the HTML and immediately opens it with Start-Process. Implies -OutputHtml.

    .PARAMETER DarkMode
    Switch. Uses the dark-themed CSS instead of the default light/print-friendly theme.

    .EXAMPLE
    # Report for a single ticket by ID, printed to console
    Get-ATTicketReport -ids 12345

    .EXAMPLE
    # Report for multiple ticket numbers, including external notes
    Get-ATTicketReport -TicketNumbers 'T20260101.0001','T20260315.0042' -IncludeExternalNotes

    .EXAMPLE
    # Suppress console output and export everything to CSV
    $r = Get-ATTicketReport -ids 12345, 67890 -SuppressConsoleOutput
    $r | ForEach-Object {
        $_.TimeEntries | Export-Csv ".\Ticket_$($_.Header.TicketNumber)_TimeEntries.csv" -NoTypeInformation
        $_.Notes       | Export-Csv ".\Ticket_$($_.Header.TicketNumber)_Notes.csv"       -NoTypeInformation
    }

    .EXAMPLE
    # Pipe ticket numbers from another command
    Get-ATTickets -TitleBeginsWith 'RMM' -LastxWeeks 4 |
        Select-Object -ExpandProperty TicketNumber |
        ForEach-Object { Get-ATTicketReport -TicketNumbers $_ }

    .EXAMPLE
    # Open HTML report in browser (light mode by default, print-friendly)
    Get-ATTicketReport -ids 12345 -OpenInBrowser

    .EXAMPLE
    # Dark-themed HTML saved to a specific path, console suppressed
    Get-ATTicketReport -TicketNumbers 'T20260101.0001' -OutputHtml -DarkMode -HtmlPath 'C:\Reports\t.html' -SuppressConsoleOutput

    .EXAMPLE
    # Multiple tickets: console + browser, external notes included
    Get-ATTicketReport -ids 12345,67890 -IncludeExternalNotes -OpenInBrowser

    .NOTES
    Ticket notes are stored in the Autotask TicketNotes entity.
    noteType values: 1 = Internal (staff only), 2 = External (client-visible).

    Time entries are fetched via Get-ATTimeEntries -ForTicketID, which requests summary notes
    automatically. Entries with no startDateTime/endDateTime are flagged HasTimestamp = $false.

    The Autotask description field on a ticket is stored as HTML. This function strips the
    most common HTML tags to produce readable plain text for console display and the Header
    object. Full raw HTML is preserved in Header.DescriptionRaw if downstream rendering
    is needed.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByID')]
    param (
        [Parameter(ParameterSetName = 'ByID', Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [int[]]$ids,

        [Parameter(ParameterSetName = 'ByNumber', Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [string[]]$TicketNumbers,

        [switch]$IncludeExternalNotes  = $false,
        [switch]$ExcludeInternalNotes  = $false,
        [switch]$SuppressConsoleOutput = $false,

        # HTML output — light mode is default (print-friendly, minimal ink)
        [switch]$OutputHtml   = $false,
        [string]$HtmlPath     = '',
        [switch]$OpenInBrowser = $false,
        # When set, use the dark-themed version instead of the default light/print mode
        [switch]$DarkMode     = $false
    )

    # ── Helper: strip common HTML tags for readable plain-text display ────────
    function ConvertFrom-ATHtml {
        param([string]$html)
        if (-not $html) { return '' }
        $text = $html `
            -replace '(?s)<br\s*/?>'          , "`n" `
            -replace '(?s)<p[^>]*>'           , "`n" `
            -replace '(?s)</p>'               , '' `
            -replace '(?s)<li[^>]*>'          , "`n  • " `
            -replace '(?s)<[^>]+>'            , '' `
            -replace '&amp;'                  , '&' `
            -replace '&lt;'                   , '<' `
            -replace '&gt;'                   , '>' `
            -replace '&nbsp;'                 , ' ' `
            -replace '&quot;'                 , '"' `
            -replace '&#39;'                  , "'" `
            -replace '(?m)^\s+$'              , '' `
            -replace "`n`n`n+"                , "`n`n"
        return $text.Trim()
    }

    # ── Helper: write a section header to console ─────────────────────────────
    function Write-ATSection {
        param([string]$Title, [string]$Color = 'Cyan')
        if (-not $SuppressConsoleOutput) {
            Write-Host ""
            Write-Host "  ── $Title " -ForegroundColor $Color -NoNewline
            Write-Host ('─' * [Math]::Max(2, 60 - $Title.Length)) -ForegroundColor DarkGray
        }
    }

    # ── Helper: build self-contained HTML report ─────────────────────────────
    function Build-ATTicketHtml {
        param(
            [System.Collections.Generic.List[PSCustomObject]]$Reports,
            [bool]$Dark = $false
        )

        # ── Utility: HTML-escape a string ─────────────────────────────────────
        function hesc {
            param([string]$s)
            if (-not $s) { return '' }
            $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
        }

        # ── Utility: use raw HTML when available, else wrap plain text ────────
        function bodyHtml {
            param([string]$raw, [string]$plain)
            if ($raw -and $raw -match '<[a-zA-Z]') { return "<div class='at-html'>$raw</div>" }
            if ($plain) { return "<pre class='pre-wrap'>$(hesc $plain)</pre>" }
            return ''
        }

        # ── CSS — two themes ──────────────────────────────────────────────────
        $darkCss = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",Arial,sans-serif;font-size:14px;background:#1a1a2e;color:#e0e0e0;padding:24px}
h1{font-size:1.4rem;color:#c8d6e5;margin-bottom:4px}
.report-meta{font-size:.8rem;color:#666;margin-bottom:28px}
.ticket{background:#16213e;border:1px solid #0f3460;border-radius:8px;margin-bottom:32px;overflow:hidden}
.t-head{background:#0f3460;padding:14px 18px}
.t-head.closed{background:#1a1a2e;border-bottom:1px solid #0f3460}
.t-num{font-size:.75rem;color:#7eb8f7;text-transform:uppercase;letter-spacing:.05em;margin-bottom:2px}
.t-title{font-size:1.1rem;font-weight:600;color:#fff}
.t-title.closed{color:#888}
.t-company{font-size:.85rem;color:#a0b4c8;margin-top:3px}
.t-meta{display:flex;flex-wrap:wrap;border-bottom:1px solid #0f3460}
.mi{padding:8px 18px;border-right:1px solid #0f3460;min-width:140px}
.mi:last-child{border-right:none}
.ml{font-size:.7rem;color:#7eb8f7;text-transform:uppercase;letter-spacing:.05em;margin-bottom:2px}
.mv{color:#e0e0e0}
.mv.ok{color:#4ade80}
.mv.warn{color:#f7c948;font-weight:600}
.mv.nb{color:#f87171}
.mv.dim{color:#888}
.hrs-bar{display:flex;flex-wrap:wrap;border-bottom:1px solid #0f3460;font-size:.82rem}
.hc{padding:7px 18px;border-right:1px solid #0f3460}
.hc:last-child{border-right:none}
.hl{font-size:.7rem;color:#7eb8f7;text-transform:uppercase}
.hv{font-weight:600;color:#e0e0e0}
.hv.ok{color:#4ade80}
.hv.warn{color:#f7c948}
.hv.nb{color:#f87171}
.sec{padding:12px 18px 0}
.sec-title{font-size:.72rem;text-transform:uppercase;letter-spacing:.08em;color:#7eb8f7;border-bottom:1px solid #0f3460;padding-bottom:5px;margin-bottom:10px}
.pre-wrap{white-space:pre-wrap;color:#c8d6e5;line-height:1.6;font-size:.88rem;padding-bottom:10px;font-family:inherit}
.at-html{color:#c8d6e5;line-height:1.6;font-size:.88rem;padding-bottom:10px}
table.eng{width:100%;border-collapse:collapse;font-size:.85rem;margin-bottom:12px}
table.eng th{text-align:left;color:#7eb8f7;font-weight:500;padding:4px 10px 4px 0;border-bottom:1px solid #0f3460}
table.eng td{padding:5px 10px 5px 0;color:#c8d6e5;border-bottom:1px solid #0a1628}
table.eng td.r{text-align:left;padding-left:14px}
table.eng td.ah{color:#f7c948}
table.eng td.nb{color:#f87171}
.note{border-left:3px solid #0f3460;margin-bottom:10px;padding:8px 12px;background:#0d1b33;border-radius:0 4px 4px 0}
.note.int{border-left-color:#f7c948}
.note.ext{border-left-color:#4ade80}
.n-meta{font-size:.75rem;color:#7eb8f7;margin-bottom:3px}
.n-meta.int .n-type{color:#f7c948;font-weight:600}
.n-meta.ext .n-type{color:#4ade80;font-weight:600}
.n-title{font-weight:600;color:#c8d6e5;margin-bottom:4px;font-size:.87rem}
.n-body{color:#b0bec5;line-height:1.55;font-size:.87rem}
.day-header{font-size:.8rem;font-weight:600;color:#7eb8f7;margin:10px 0 4px;padding:3px 0;border-bottom:1px solid #0f3460}
.te{background:#0d1b33;border-radius:4px;margin-bottom:6px;padding:7px 12px}
.te-hdr{display:flex;flex-wrap:wrap;gap:10px;align-items:baseline;margin-bottom:3px}
.te-time{color:#7eb8f7;font-size:.85rem}
.te-hrs{font-weight:600;color:#4ade80}
.te-hrs.nb{color:#f87171}
.te-hrs.ah{color:#f7c948}
.te-eng{color:#a0b4c8;font-size:.85rem}
.te-bill{font-size:.76rem;color:#555;margin-bottom:2px}
.te-notes{color:#b0bec5;white-space:pre-wrap;font-size:.86rem;line-height:1.5}
.te-warn{font-size:.76rem;color:#f7c948;margin-top:2px}
.no-stamp{background:#1a1500;border:1px solid #f7c948;border-radius:4px;padding:5px 12px;font-size:.8rem;color:#f7c948;margin-top:6px}
.summ-box{background:#16213e;border:1px solid #0f3460;border-radius:8px;padding:14px 18px;margin-bottom:24px}
.summ-box h2{font-size:.8rem;text-transform:uppercase;letter-spacing:.08em;color:#7eb8f7;margin-bottom:10px}
table.summ{width:100%;border-collapse:collapse;font-size:.84rem}
table.summ th{text-align:left;color:#7eb8f7;font-weight:500;padding:4px 10px 5px 0;border-bottom:1px solid #0f3460}
table.summ td{padding:5px 10px 5px 0;border-bottom:1px solid #0a1628;color:#c8d6e5}
table.summ td.r{text-align:right;padding-right:18px}
table.summ td.ok{color:#4ade80}
table.summ td.nb{color:#f87171}
.grand{display:flex;gap:24px;margin-top:8px;padding-top:8px;border-top:1px solid #0f3460;font-size:.88rem}
.gi .gl{font-size:.7rem;color:#7eb8f7;text-transform:uppercase}
.gi .gv{font-weight:600}
.gi .gv.ok{color:#4ade80}
.gi .gv.nb{color:#f87171}
'@

        $lightCss = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",Arial,sans-serif;font-size:13px;background:#fff;color:#111;padding:20px}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:3px}
.report-meta{font-size:.78rem;color:#777;margin-bottom:22px}
@media print{body{padding:8px}.ticket{page-break-inside:avoid;margin-bottom:18px}}
<!-- 
@media print {
      .no-page-break {
        break-inside: avoid;
        page-break-inside: avoid;
      }

      h1, h2, h3 {
        break-after: avoid;
        page-break-after: avoid;
      }
    }
-->

.ticket{border:1px solid #bbb;border-radius:6px;margin-bottom:24px;overflow:hidden}
.t-head{background:#e8edf2;padding:10px 16px;border-bottom:1px solid #bbb}
.t-head.closed{background:#f5f5f5}
.t-num{font-size:1rem;color:#3a5a8c;text-transform:uppercase;letter-spacing:.05em;margin-bottom:1px}
.t-title{font-size:1.05rem;font-weight:700;color:#111}
.t-title.closed{color:#666}
.t-company{font-size:.83rem;color:#444;margin-top:2px}
.t-meta{display:flex;flex-wrap:wrap;border-bottom:1px solid #ccc}
.mi{padding:6px 14px;border-right:1px solid #ddd;min-width:130px}
.mi:last-child{border-right:none}
.ml{font-size:.68rem;color:#3a5a8c;text-transform:uppercase;letter-spacing:.04em;margin-bottom:1px}
.mv{color:#111}
.mv.ok{color:#1a7a3a}
.mv.warn{color:#b06000;font-weight:700}
.mv.nb{color:#c0000a}
.mv.dim{color:#777}
.hrs-bar{display:flex;flex-wrap:wrap;border-bottom:1px solid #ccc;font-size:.8rem}
.hc{padding:5px 14px;border-right:1px solid #ddd}
.hc:last-child{border-right:none}
.hl{font-size:.67rem;color:#3a5a8c;text-transform:uppercase}
.hv{font-weight:700;color:#111}
.hv.ok{color:#1a7a3a}
.hv.warn{color:#b06000}
.hv.nb{color:#c0000a}
.sec{padding:10px 16px 0}
.sec-title{font-size:.69rem;text-transform:uppercase;letter-spacing:.07em;color:#3a5a8c;border-bottom:1px solid #ccc;padding-bottom:4px;margin-bottom:8px}
.pre-wrap{white-space:pre-wrap;color:#222;line-height:1.55;font-size:.87rem;padding-bottom:8px;font-family:inherit}
.at-html{color:#222;line-height:1.55;font-size:.87rem;padding-bottom:8px}
table.eng{width:100%;border-collapse:collapse;font-size:.83rem;margin-bottom:10px}
table.eng th{text-align:left;color:#3a5a8c;font-weight:600;padding:3px 10px 3px 0;border-bottom:1px solid #ccc}
table.eng td{padding:4px 10px 4px 0;color:#111;border-bottom:1px solid #eee}
table.eng td.r{text-align:left;padding-left:14px}
table.eng td.ah{color:#b06000}
table.eng td.nb{color:#c0000a}
.note{border-left:3px solid #bbb;margin-bottom:8px;padding:6px 10px;background:#fafafa;border-radius:0 3px 3px 0}
.note.int{border-left-color:#c07000;background:#fffbf0}
.note.ext{border-left-color:#1a7a3a;background:#f0fff4}
.n-meta{font-size:.72rem;color:#555;margin-bottom:2px}
.n-meta.int .n-type{color:#c07000;font-weight:700}
.n-meta.ext .n-type{color:#1a7a3a;font-weight:700}
.n-title{font-weight:700;color:#222;margin-bottom:3px;font-size:.86rem}
.n-body{color:#333;line-height:1.5;font-size:.86rem}
.day-header{font-size:.78rem;font-weight:700;color:#3a5a8c;margin:8px 0 3px;padding:2px 0;border-bottom:1px solid #ddd}
.te{border:1px solid #eee;border-radius:3px;margin-bottom:5px;padding:6px 10px;background:#fdfdfd}
.te-hdr{display:flex;flex-wrap:wrap;gap:8px;align-items:baseline;margin-bottom:2px}
.te-time{color:#3a5a8c;font-size:.83rem}
.te-hrs{font-weight:700;color:#1a7a3a}
.te-hrs.nb{color:#c0000a}
.te-hrs.ah{color:#b06000}
.te-eng{color:#444;font-size:.83rem}
.te-bill{font-size:.74rem;color:#888;margin-bottom:2px}
.te-notes{color:#333;white-space:pre-wrap;font-size:.85rem;line-height:1.45}
.te-warn{font-size:.74rem;color:#b06000;margin-top:2px}
.no-stamp{border:1px solid #b06000;border-radius:3px;padding:4px 10px;font-size:.78rem;color:#b06000;margin-top:5px}
.summ-box{border:1px solid #bbb;border-radius:6px;padding:12px 16px;margin-bottom:20px}
.summ-box h2{font-size:.78rem;text-transform:uppercase;letter-spacing:.07em;color:#3a5a8c;margin-bottom:8px}
table.summ{width:100%;border-collapse:collapse;font-size:.82rem}
table.summ th{text-align:left;color:#3a5a8c;font-weight:600;padding:3px 10px 4px 0;border-bottom:1px solid #ccc}
table.summ td{padding:5px 10px 5px 0;border-bottom:1px solid #eee;color:#111}
table.summ td.r{text-align:right;padding-right:18px}
table.summ td.ok{color:#1a7a3a}
table.summ td.nb{color:#c0000a}
.grand{display:flex;gap:22px;margin-top:7px;padding-top:7px;border-top:1px solid #ccc;font-size:.86rem}
.gi .gl{font-size:.68rem;color:#3a5a8c;text-transform:uppercase}
.gi .gv{font-weight:700}
.gi .gv.ok{color:#1a7a3a}
.gi .gv.nb{color:#c0000a}
'@

        $css = if ($Dark) { $darkCss } else { $lightCss }

        # ── Build per-ticket blocks ────────────────────────────────────────────
        $blocks = [System.Text.StringBuilder]::new()
        $gen   = (Get-Date).ToString('yyyy-MM-dd HH:mm')

        foreach ($rpt in $Reports) {
            $hdr  = $rpt.Header
            $nts  = $rpt.Notes
            $tes  = $rpt.TimeEntries
            $summ = $rpt.Summary
            $engs = $rpt.EngineerSummary

            $isOpen  = -not $hdr.CompletedDate
            $hclass  = if ($isOpen) { 'ticket-header t-head' } else { 't-head closed' }
            $ttclass = if ($isOpen) { 't-title' } else { 't-title closed' }

            $billCls = if ($hdr.EstimatedHrs -gt 0 -and $summ.BillableHours -gt $hdr.EstimatedHrs) { 'warn' } else { 'ok' }

            $null = $blocks.AppendLine("<div class='ticket'>")

            # Header
            $null = $blocks.AppendLine("  <div class='$hclass'>")
            $null = $blocks.AppendLine("    <div class='t-num'>$(hesc $hdr.TicketNumber)</div>")
            $null = $blocks.AppendLine("    <div class='$ttclass'>$(hesc $hdr.Title)</div>")
            if ($hdr.Company) { $null = $blocks.AppendLine("    <div class='t-company'>$(hesc $hdr.Company)</div>") }
            $null = $blocks.AppendLine("  <div class='t-company'>Report generated on $gen</div>")
            $null = $blocks.AppendLine("  </div>")

            # Meta row
            $null = $blocks.AppendLine("  <div class='t-meta'>")
            $sCls = if ($hdr.CompletedDate) { 'dim' } else { 'ok' }
            $null = $blocks.AppendLine("    <div class='mi'><div class='ml'>Status</div><div class='mv $sCls'>$(hesc $hdr.Status)</div></div>")
            if ($hdr.Queue) { $null = $blocks.AppendLine("    <div class='mi'><div class='ml'>Queue</div><div class='mv'>$(hesc $hdr.Queue)</div></div>") }
            $null = $blocks.AppendLine("    <div class='mi'><div class='ml'>Assigned To</div><div class='mv'>$(hesc $hdr.AssignedTo)</div></div>")
            $null = $blocks.AppendLine("    <div class='mi'><div class='ml'>Created</div><div class='mv'>$($hdr.CreateDate -replace 'T.*','')</div></div>")
            if ($hdr.CompletedDate) {
                $null = $blocks.AppendLine("    <div class='mi'><div class='ml'>Completed</div><div class='mv dim'>$($hdr.CompletedDate -replace 'T.*','')</div></div>")
            }
            $null = $blocks.AppendLine("  </div>")

            # Hours bar
            $null = $blocks.AppendLine("  <div class='hrs-bar'>")
            $null = $blocks.AppendLine("    <div class='hc'><div class='hl'>Billable</div><div class='hv $billCls'>$($summ.BillableHours) h</div></div>")
            if ($summ.NonBillableHours -gt 0) {
                $null = $blocks.AppendLine("    <div class='hc'><div class='hl'>Non-Billable</div><div class='hv nb'>$($summ.NonBillableHours) h</div></div>")
            }
            $null = $blocks.AppendLine("    <div class='hc'><div class='hl'>Total</div><div class='hv'>$($summ.TotalHours) h</div></div>")
            if ($hdr.EstimatedHrs -gt 0) {
                $null = $blocks.AppendLine("    <div class='hc'><div class='hl'>Estimated</div><div class='hv'>$($hdr.EstimatedHrs) h</div></div>")
            }
            if ($summ.AfterHours -gt 0) {
                $null = $blocks.AppendLine("    <div class='hc'><div class='hl'>After Hours</div><div class='hv warn'>$($summ.AfterHours) h</div></div>")
            }
            if ($summ.DurationOnlyHours -gt 0) {
                $null = $blocks.AppendLine("    <div class='hc'><div class='hl'>No Timestamp</div><div class='hv warn'>$($summ.DurationOnlyHours) h &#9888;</div></div>")
            }
            $null = $blocks.AppendLine("  </div>")

            # Engineer effort table — th text-align:left is enforced by .eng th in both CSS themes
            if ($engs -and $engs.Count -gt 0) {
                $null = $blocks.AppendLine("  <div class='sec'>")
                $null = $blocks.AppendLine("    <div class='sec-title'>Engineer Effort</div>")
                $null = $blocks.AppendLine("    <table class='eng'><thead><tr>")
                $null = $blocks.AppendLine("      <th>Engineer</th><th class='r'>Hours</th><th class='r'>After Hrs</th><th class='r'>Non-Billable</th>")
                $null = $blocks.AppendLine("    </tr></thead><tbody>")
                foreach ($eng in $engs) {
                    $ahc = if ($eng.AfterHours   -gt 0) { 'ah' } else { '' }
                    $nbc = if ($eng.NonBillableHrs -gt 0) { 'nb' } else { '' }
                    $ahv = if ($eng.AfterHours   -gt 0) { $eng.AfterHours }    else { '-' }
                    $nbv = if ($eng.NonBillableHrs -gt 0) { $eng.NonBillableHrs } else { '-' }
                    $null = $blocks.AppendLine("    <tr><td>$(hesc $eng.Name)</td><td class='r'>$($eng.Hours)</td><td class='r $ahc'>$ahv</td><td class='r $nbc'>$nbv</td></tr>")
                }
                $null = $blocks.AppendLine("    </tbody></table></div>")
            }

            # Description
            $dh = bodyHtml -raw $($hdr.DescriptionRaw -replace "(\r?\n){2,}","\r\n") -plain $($hdr.Description -replace "(\r?\n){2,}", "\r\n")
            if ($dh) {
                $null = $blocks.AppendLine("  <div class='sec'><div class='sec-title'>Description</div>$dh</div>")
            }

            # Notes
            if ($nts -and $nts.Count -gt 0) {
                $null = $blocks.AppendLine("  <div class='sec'>")
                $null = $blocks.AppendLine("    <div class='sec-title'>Notes ($($nts.Count))</div>")
                foreach ($n in $nts) {
                    $nc  = $n.NoteType.ToLower().Substring(0,3)   # 'int' or 'ext'
                    $ds  = $n.CreatedDate -replace 'T',' ' -replace '\.\d+',''
                    $null = $blocks.AppendLine("    <div class='note $nc'>")
                    $null = $blocks.AppendLine("      <div class='n-meta $nc'><span class='n-type'>$(hesc $n.NoteType)</span>  $ds &mdash; $(hesc $n.Author)</div>")
                    if ($n.Title) { $null = $blocks.AppendLine("      <div class='n-title'>$(hesc $n.Title)</div>") }
                    $null = $blocks.AppendLine("      <div class='n-body'>$(bodyHtml -raw $($n.NoteTextRaw -replace "(\r?\n){2,}", "\r\n") -plain $($n.NoteText -replace "(\r?\n){2,}", "\r\n"))</div>")
                    $null = $blocks.AppendLine("    </div>")
                }
                $null = $blocks.AppendLine("  </div>")
            }

            # Time entries — grouped by day
            $null = $blocks.AppendLine("  <div class='sec'>")
            $null = $blocks.AppendLine("    <div class='sec-title'>Time Entries ($($tes.Count)) &mdash; Total: $($summ.TotalHours) h &mdash; Billable: $($summ.BillableHours) h</div>")
            if ($tes -and $tes.Count -gt 0) {
                $lastDay = $null
                foreach ($te in $tes) {
                    # Day separator
                    if ($te.DateWorked -ne $lastDay) {
                        $null = $blocks.AppendLine("    <div class='day-header'>&#9656; $(([datetime]$te.DateWorked).ToShortDateString())</div>")
                        $lastDay = $te.DateWorked
                    }
                    $hrsClass = if ($te.AfterHours -gt 0) { 'ah' } elseif (-not $te.IsBillable) { 'nb' } else { '' }
                    $timeStr  = if ($te.HasTimestamp) { "$($te.StartTime)&ndash;$($te.EndTime)" } else { '' }
                    $tags     = @()
                    if (-not $te.IsBillable) { $tags += '[NB]' }
                    if ($te.AfterHours -gt 0) { $tags += '[AH]' }
                    $tagStr   = if ($tags) { ' ' + ($tags -join ' ') } else { '' }
                    $null = $blocks.AppendLine("    <div class='te'>")
                    $null = $blocks.AppendLine("      <div class='te-hdr'>")
                    if ($timeStr) { $null = $blocks.AppendLine("        <span class='te-time'>$timeStr</span>") }
                    $null = $blocks.AppendLine("        <span class='te-hrs $hrsClass'>$($te.HoursWorked) h$tagStr</span>")
                    $null = $blocks.AppendLine("        <span class='te-eng'>$(hesc $te.Engineer)</span>")
                    $null = $blocks.AppendLine("      </div>")
                    if ($te.BillingCode) { $null = $blocks.AppendLine("      <div class='te-bill'>$(hesc $te.BillingCode)</div>") }
                    if ($te.SummaryNotes) {
                        $null = $blocks.AppendLine("      <div class='te-notes'>$(bodyHtml -raw '' -plain $($te.SummaryNotes -replace "(\r?\n){2,}", "\r\n"))</div>")
                    }
                    if (-not $te.HasTimestamp) {
                        $null = $blocks.AppendLine("      <div class='te-warn'>&#9888; Duration-only &mdash; no clock start/end recorded</div>")
                    }
                    $null = $blocks.AppendLine("    </div>")
                }
                if ($summ.DurationOnlyHours -gt 0) {
                    $null = $blocks.AppendLine("    <div class='no-stamp'>&#9888; $($summ.DurationOnlyHours) h recorded without clock start/end time</div>")
                }
            } else {
                $null = $blocks.AppendLine("    <p style='color:#999;padding:6px 0'>(no time entries)</p>")
            }
            $null = $blocks.AppendLine("  </div><div style='height:10px'></div>")
            $null = $blocks.AppendLine("</div>")
        }

        # Summary table (multiple tickets)
        $summHtml = ''
        if ($Reports.Count -gt 1) {
            $all   = $Reports | ForEach-Object { $_.Summary }
            $gTot  = [Math]::Round(($all | Measure-Object TotalHours     -Sum).Sum, 2)
            $gBill = [Math]::Round(($all | Measure-Object BillableHours  -Sum).Sum, 2)
            $gNB   = [Math]::Round(($all | Measure-Object NonBillableHours -Sum).Sum, 2)
            $rows  = [System.Text.StringBuilder]::new()
            foreach ($s in ($all | Sort-Object TicketNumber)) {
                $bc = if ($s.BillableHours -gt 0) { 'ok' } else { '' }
                $nc = if ($s.NonBillableHours -gt 0) { 'nb' } else { '' }
                $nv = if ($s.NonBillableHours -gt 0) { $s.NonBillableHours } else { '-' }
                $null = $rows.AppendLine("<tr><td>$(hesc $s.TicketNumber)</td><td>$(hesc ($s.Title -replace '(.{55}).+','$1\u2026'))</td><td>$(hesc $s.Status)</td><td>$(hesc $s.AssignedTo)</td><td class='r $bc'>$($s.BillableHours)</td><td class='r $nc'>$nv</td><td class='r'>$($s.TotalHours)</td></tr>")
            }
            $summHtml = "<div class='summ-box'><h2>Report Summary &mdash; $($Reports.Count) Tickets</h2><table class='summ'><thead><tr><th>Ticket</th><th>Title</th><th>Status</th><th>Assigned</th><th class='r'>Billable h</th><th class='r'>Non-Bill h</th><th class='r'>Total h</th></tr></thead><tbody>$($rows.ToString())</tbody></table><div class='grand'><div class='gi'><div class='gl'>Grand Total</div><div class='gv'>$gTot h</div></div><div class='gi'><div class='gl'>Billable</div><div class='gv ok'>$gBill h</div></div><div class='gi'><div class='gl'>Non-Billable</div><div class='gv nb'>$gNB h</div></div></div></div>"
        }

        $theme = if ($Dark) { 'dark' } else { 'light (print-friendly)' }
        $title = if ($Reports.Count -eq 1) { "Ticket Report &mdash; $(hesc $Reports[0].Header.TicketNumber)" } else { "Ticket Report &mdash; $($Reports.Count) tickets" }
      #  $gen   = (Get-Date).ToString('yyyy-MM-dd HH:mm')

        return @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>$title</title>
  <style>$css
  </style>
  


</head>
<body>
<!--
  <h1>$title</h1>
 
  <div class='report-meta'>Generated $gen &mdash; AutoTaskRest &mdash; Theme: $theme</div>
  -->
  $summHtml
  $($blocks.ToString())
</body>
</html>
"@
    }

    # ── 1. Resolve ticket IDs ─────────────────────────────────────────────────
    $resolvedIDs = [System.Collections.Generic.List[int]]::new()

    if ($PSCmdlet.ParameterSetName -eq 'ByNumber') {
        Write-Host "Get-ATTicketReport: resolving $($TicketNumbers.Count) ticket number(s)..." -ForegroundColor Cyan
        $found = Get-ATTickets -TicketNumbers $TicketNumbers -GetCompanyNames -ReturnAllFields 
        if (-not $found) {
            Write-Warning "Get-ATTicketReport: no tickets found matching the supplied ticket numbers."
            return
        }
        foreach ($t in $found) { $resolvedIDs.Add($t.id) }
    }
    else {
        $found = Get-ATTickets -ids $ids  -ReturnAllFields -GetCompanyNames
        foreach ($i in $found) { $resolvedIDs.Add($i.id) }
    }

    if ($resolvedIDs.Count -eq 0) {
        Write-Warning "Get-ATTicketReport: no ticket IDs to process."
        return
    }

    Write-Host "Get-ATTicketReport: fetching report for $($resolvedIDs.Count) ticket(s)..." -ForegroundColor Cyan

    # $engIDs = $found | Select-Object -ExpandProperty assignedResourceID -Unique

    # ── 2. Pre-fetch lookup tables (once, shared across all tickets) ──────────
    $ticketFieldInfo = Get-ATTicketFieldInfo
    $engineers = Get-ATEngineers -id $found.assignedResourceID
    $billingCodes = Get-ATBillingCodes

    $noteTypeFilter = @()
    if (!$ExcludeInternalNotes) { $noteTypeFilter += '{"op":"eq","Field":"noteType","value":"1"}' }
    if ($IncludeExternalNotes) { $noteTypeFilter += '{"op":"eq","Field":"noteType","value":"2"}' }
    if ($noteTypeFilter.Count -gt 0) {
        $noteTypeClause = if ($noteTypeFilter.Count -eq 1) {
            $noteTypeFilter[0]
        }
        else {
            '{"op":"or","items":[' + ($noteTypeFilter -join ',') + ']}'
        }
        $ResolvedIDstr = $resolvedIDs -join ','
        $notesFilter = '{"op":"and","items":[{"op":"eq","Field":"ticketID","value":"' + $ResolvedIDstr + '"},' + $noteTypeClause + ']}'
        $allNotes = Invoke-ATQuery -entityName 'v1.0/TicketNotes' `
            -SearchFirstBy Nothing `
            -SearchFurtherBy $notesFilter

        # ── 3. Build a report object per ticket ───────────────────────────────────
        $reports = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($ticketID in $resolvedIDs) {

            # ── 3a. Fetch full ticket header (all fields) ─────────────────────────
            Write-Host "  Fetching ticket ID $ticketID..." -ForegroundColor DarkCyan
            $rawTicket = $found | Where-Object id -eq $ticketID | Select-Object -First 1


            if (-not $rawTicket) {
                Write-Warning "  Ticket $ticketID not found — skipping."
                continue
            }

            # Resolve human-readable field values from picklists
            $statusName = ($ticketFieldInfo.status   | Where-Object value -eq $rawTicket.status   | Select-Object -First 1).label
            $queueName = ($ticketFieldInfo.queueID   | Where-Object value -eq $rawTicket.queueID  | Select-Object -First 1).label
            $assignedName = if ($rawTicket.assignedResourceID) {
                ($engineers | Where-Object id -eq $rawTicket.assignedResourceID | Select-Object -First 1).FullName
            }
            else { '' }

            $descPlain = (ConvertFrom-ATHtml -html $rawTicket.description) -replace "(\r?\n){2,}", "\r\n"

            $header = [PSCustomObject]@{
                TicketID       = $rawTicket.id
                TicketNumber   = $rawTicket.ticketNumber
                Company        = $rawTicket.CompanyName
                Title          = $rawTicket.title
                Description    = $descPlain -replace "(\r?\n){2,}", "\r\n"
                DescriptionRaw = $rawTicket.description -replace "(\r?\n){2,}", "`n"
                Status         = $statusName
                Queue          = $queueName
                AssignedTo     = $assignedName
                CompanyID      = $rawTicket.companyID
                CreateDate     = $rawTicket.createDate
                LastActivity   = $rawTicket.lastActivityDate
                CompletedDate  = $rawTicket.completedDate
                BillingCodeID  = $rawTicket.billingCodeID
                IsInternal     = ($rawTicket.companyID -in @(0, 1))   # refined below once company is known
                EstimatedHrs   = $rawTicket.estimatedHours
            }

         #   write-host "  Ticket1: $($header.TicketNumber) - $($header.EstimatedHrs)" -ForegroundColor Green

            # ── 3b. Fetch all ticket notes ────────────────────────────────────────

            $noteRows = [System.Collections.Generic.List[PSCustomObject]]::new()



            $rawnotes = $allNotes | Where-Object { $_.ticketID -eq $ticketID }
            
            # $rawNotes = Invoke-ATQuery -entityName 'v1.0/TicketNotes' `
            #     -SearchFirstBy Nothing `
            #     -SearchFurtherBy $notesFilter

            if ($rawNotes) {
                foreach ($n in ($rawNotes | Sort-Object createDateTime)) {
                    $authorName = if ($n.creatorResourceID) {
                        ($engineers | Where-Object id -eq $n.creatorResourceID | Select-Object -First 1).FullName
                    }
                    else { '' }

                    $noteRows.Add([PSCustomObject]@{
                            TicketID     = $ticketID
                            TicketNumber = $header.TicketNumber
                            Company      = $header.Company
                            NoteID       = $n.id
                            NoteType     = if ($n.noteType -eq 1) { 'Internal' } else { 'External' }
                            Title        = $n.title
                            NoteText     = ConvertFrom-ATHtml -html $n.description #-replace "(\r?\n){2,}", "`n"
                            NoteTextRaw  = ($n.description -replace "(\r?\n){2,}", "`n")
                            Author       = $authorName
                            CreatedDate  = $n.createDateTime
                            LastModified = $n.lastModifiedDateTime
                        })
                }
            }
        }

        # ── 3c. Fetch all time entries for this ticket ────────────────────────
        $rawEntries = Get-ATTimeEntries `
            -ForTicketID      $ticketID `
            -IncludeSummaryNotes   
        #-IncludeBillingDetails 
        #-includeEngineerDetails

        $entryRows = [System.Collections.Generic.List[PSCustomObject]]::new()
      #      write-host "  Ticket2: $($header.TicketNumber) - $($header.EstimatedHrs)" -ForegroundColor Green

        if ($rawEntries) {
            foreach ($e in ($rawEntries | Sort-Object dateWorked)) {
                $hasTimestamp = ($null -ne $e.startDateTime) -and ($null -ne $e.endDateTime)
                $engineerLabel = if ($e.Engineer) { $e.Engineer } else { "ID:$($e.resourceID)" }
                $bcName = if ($e.BillingCode) { $e.BillingCode } `
                    else { ($billingCodes | Where-Object id -eq $e.billingCodeID | Select-Object -First 1).name }

                $startDisplay = if ($hasTimestamp -and $e.startDateTimeLocal) {
                    $e.startDateTimeLocal.ToString('HH:mm')
                }
                else { '' }
                $endDisplay = if ($hasTimestamp -and $e.endDateTimeLocal) {
                    $e.endDateTimeLocal.ToString('HH:mm')
                }
                else { '' }

            #    Write-Host "  Billable Hours        : $billableHrs       Total hrs: $totalHrs" -ForegroundColor $headerColor


                $entryRows.Add([PSCustomObject]@{
                        TicketID           = $ticketID
                        TicketNumber       = $header.TicketNumber
                        TimeEntryID        = $e.id
                        DateWorked         = ($e.dateWorked -replace 'T.*', '')
                        StartTime          = $startDisplay
                        EndTime            = $endDisplay
                        HasTimestamp       = $hasTimestamp
                        Engineer           = $engineerLabel
                        HoursWorked        = [Math]::Round($e.hoursWorked, 2)
                        AfterHours         = $e.afterHours
                        ClassificationIcon = $e.CompanyClassification
                        IsBillable         = -not ($e.isNonBillable -eq $true)
                        BillingCode        = $bcName
                        SummaryNotes       = if ($e.summaryNotes) { $e.summaryNotes } else { '' }
                        NonBillableHrs     = $e.hoursNonBillable
                        
                    })
            }
        }

                #    write-host "  Ticket3: $($header.TicketNumber) - $($header.EstimatedHrs)" -ForegroundColor Green

        # ── 3d. Build per-ticket summary numbers ──────────────────────────────
        $totalHrs = [Math]::Round(($entryRows | Measure-Object HoursWorked -Sum).Sum, 2)
        $billableHrs = [Math]::Round(($entryRows | Where-Object IsBillable | Measure-Object HoursWorked -Sum).Sum, 2)
        $noStampHrs = [Math]::Round(($entryRows | Where-Object { -not $_.HasTimestamp } | Measure-Object HoursWorked -Sum).Sum, 2)
        $afterHrs = [Math]::Round(($entryRows | Measure-Object AfterHours -Sum).Sum, 2)
        $NonBillableHrs = [Math]::Round(($entryRows  | Measure-Object NonBillableHrs -Sum).Sum, 2)

        $ticketSummary = [PSCustomObject]@{
            TicketID          = $header.TicketID
            TicketNumber      = $header.TicketNumber
            Company           = $header.CompanyID
            Title             = $header.Title
            Status            = $header.Status
            AssignedTo        = $header.AssignedTo
            TotalHours        = $totalHrs
            BillableHours     = $billableHrs
            NonBillableHours  = $NonBillableHrs #[Math]::Round($totalHrs - $billableHrs, 2)
            DurationOnlyHours = $noStampHrs
            NoteCount         = $noteRows.Count
            TimeEntryCount    = $entryRows.Count
            AfterHours        = $afterHrs

        }

        # ── 3e. Console report ────────────────────────────────────────────────
        if (-not $SuppressConsoleOutput) {
            $isOpen = -not $rawTicket.completedDate
            $headerColor = if ($isOpen) { 'White' } else { 'DarkGray' }

            Write-Host ""
            Write-Host ('═' * 70) -ForegroundColor DarkGray
            Write-Host "  TICKET   $($header.TicketNumber)   $($header.Title)" -ForegroundColor Yellow
            Write-Host "  Company  $($header.Company)" -ForegroundColor Yellow
            Write-Host ('═' * 70) -ForegroundColor DarkGray

            Write-Host "  Status    : $($header.Status)" -ForegroundColor $headerColor
            # Write-Host "  Queue     : $($header.Queue)"  -ForegroundColor $headerColor
            Write-Host "  Assigned  : $($header.AssignedTo)" -ForegroundColor $headerColor
            Write-Host "  Created   : $($header.CreateDate -replace 'T.*','')" -ForegroundColor $headerColor





          #  write-host "  Ticket1: $($header.TicketNumber) - $($header.EstimatedHrs)" -ForegroundColor Green

            $hrstr = "  Billable Hours        : $billableHrs"   
            if ($billableHrs -ne $totalHrs) {
                $hrstr = "  $hrstr  (Total hrs: $totalHrs) "
               # Write-Host "  Billable Hours        : $billableHrs       ( Total hrs: $totalHrs ) " -ForegroundColor $headerColor
            }
            if ($header.EstimatedHrs -gt 0){
                 $hrstr = "  $hrstr  Estimated Hrs: $($header.EstimatedHrs)"
            }
            if ($billableHrs -gt $ticketSummary.EstimatedHrs)  {
                $bColor = "Yellow"
                $hrstr = "**$hrstr**"
            }
            else {
                $bColor = $headerColor
            }
            
            Write-Host "$hrstr" -ForegroundColor $bColor

            if ($afterHrs -gt 0) {
                Write-Host "  (includes After Hours): $afterHrs" -ForegroundColor DarkYellow
            }
            if ($noStampHrs -gt 0) {
                Write-Host "  (includes $noStampHrs hrs with no clock-in/out time)" -ForegroundColor DarkYellow
            }
            If ($ticketSummary.NonBillableHours -gt 0) {
                Write-Host "  Non-Billable Hours    : $($ticketSummary.NonBillableHours)" -ForegroundColor $headerColor
            }
            if ($header.CompletedDate) {
                Write-Host "  Completed             : $($header.CompletedDate -replace 'T.*','')" -ForegroundColor DarkGray
            }
            

                #ENgineer Summar
                Write-ATSection 'Engineer effort'
                $engineers = @()
                $entryRows | Group-Object Engineer | ForEach-Object {  
                    $Hours = ($_.Group | Measure-Object HoursWorked -Sum).Sum
                    $Engineer = $_.Name
                    if ($Engineer -eq '') {
                        $Engineer = '(Unknown)'
                    }
                    $AfterHours = ($_.Group | Measure-Object AfterHours -Sum).Sum
                    $NonBillableHrs = ($_.Group | Where-Object NonBillableHrs | Measure-Object NonBillableHrs -Sum).Sum
                    $engineers += [PSCustomObject]@{
                        Name        = $Engineer
                        Hours       = $Hours
                        AfterHours  = $AfterHours
                        NonBillableHrs = $NonBillableHrs
                    }
                    Write-Host "$($Engineer.PadRight(30)): $($Hours) hrs" -ForegroundColor Gray
                } 


                foreach ($e in $engineers | Where-Object { $_.AfterHours -gt 0 }) {
                    Write-Host "  Includes After Hours         $($Engineer.PadRight(30)): $($e.AfterHours) hrs" -ForegroundColor DarkYellow
                }
                foreach ($e in $engineers | Where-Object { $_.NonBillableHrs -gt 0 }) {
                    Write-Host "  Includes Non-Billable Hours  $($Engineer.PadRight(30)): $($e.NonBillableHrs) hrs" -ForegroundColor Red
               
            

                    # $engineerHrs = $entryRows | Group-Object Engineer | Select-Object Name, @{Name='Hours';Expression={($_.Group | Measure-Object HoursWorked -Sum).Sum}}
                    # #Write-Host "  Engineer Summary:" -ForegroundColor $headerColor
                    # foreach ($eh in $engineerHrs) {
                    # }  
                } 
        

                # Description
                if ($header.Description) {
                    Write-ATSection 'Description'
                    $header.Description -split "`n" | ForEach-Object {
                        Write-Host "    $_" -ForegroundColor Gray
                    }
                }

                # Notes
                if ($noteRows.Count -gt 0) {
                    Write-ATSection "Notes ($($noteRows.Count))"
                    foreach ($note in $noteRows) {
                        $noteColor = if ($note.NoteType -eq 'Internal') { 'DarkYellow' } else { 'DarkCyan' }
                        Write-Host ""
                        Write-Host "    [$($note.NoteType)]  $($note.CreatedDate -replace 'T',' ' -replace '\.\d+','')  — $($note.Author)" -ForegroundColor $noteColor
                        if ($note.Title) {
                            Write-Host "    Subject: $($note.Title)" -ForegroundColor $noteColor
                        }
                        ($note.NoteText -replace "(\r?\n){2,}", "`n") -split "`n" | ForEach-Object {
                            Write-Host "      $_" -ForegroundColor Gray
                        }
                    }
                }
                else {
                    Write-ATSection 'Notes'
                    Write-Host "    (no notes)" -ForegroundColor DarkGray
                }

                # Time entries — grouped by day so same-day entries are visually clustered
                Write-ATSection "Time Entries ($($entryRows.Count))   Total: $totalHrs hrs   Billable: $billableHrs hrs"
                if ($entryRows.Count -gt 0) {
                    $lastDay = $null
                    foreach ($te in $entryRows) {
                        # Print a day-separator whenever the date changes
                        if ($te.DateWorked -ne $lastDay) {
                            Write-Host ""
                         #  Write-Host "    ▸ $($te.DateWorked.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
                            Write-Host "    ▸ $(([datetime]$te.DateWorked).ToShortDateString())" -ForegroundColor Cyan
                            $lastDay = $te.DateWorked
                        }
                        if ($te.AfterHours -gt 0) {$eColor = "DarkYellow"}
                        elseif ($te.NonBillableHrs -gt 0) {$eColor = "Red"}
                        else {$eColor ="DarkGray"}
                        $timeStr = if ($te.HasTimestamp) { "$($te.StartTime)–$($te.EndTime)" } else { "(no time⚠)" }
                        $billTag = if ($te.IsBillable) { '' } else { ' [NB]' }
                        Write-Host "      $timeStr  $($te.HoursWorked)h  $($te.Engineer)$billTag" -ForegroundColor White
                        Write-Host "      Billing: $($te.BillingCode)" -ForegroundColor $eColor
                        if ($te.SummaryNotes) {
                            ($te.SummaryNotes -replace "(\r?\n){2,}", "`n") -split "`n" | ForEach-Object {
                                Write-Host "        $_" -ForegroundColor Gray
                            }
                        }
                    }
                    if ($noStampHrs -gt 0) {
                        Write-Host ""
                        Write-Host "    ⚠  $noStampHrs hrs recorded without a clock start/end time" -ForegroundColor DarkYellow
                    }
                }
                else {
                    Write-Host "    (no time entries)" -ForegroundColor DarkGray
                }

                Write-Host ""
                Write-Host ('═' * 70) -ForegroundColor DarkGray
            }

            $reports.Add([PSCustomObject]@{
                    Header      = $header
                    Notes       = $noteRows.ToArray()
                    TimeEntries = $entryRows.ToArray()
                    Summary     = $ticketSummary
                    EngineerSummary = $engineers
                })
        }

        # ── 4. If multiple tickets, print a grand rollup ──────────────────────────
        if ($reports.Count -gt 1 -and -not $SuppressConsoleOutput) {
            $allSummaries = $reports | Select-Object -ExpandProperty Summary
            $grandHrs = [Math]::Round(($allSummaries | Measure-Object TotalHours    -Sum).Sum, 2)
            $grandBill = [Math]::Round(($allSummaries | Measure-Object BillableHours -Sum).Sum, 2)

            Write-Host ""
            Write-Host "─────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
            Write-Host "  REPORT SUMMARY — $($reports.Count) tickets" -ForegroundColor White
            Write-Host "  Total hours   : $grandHrs" -ForegroundColor White
            Write-Host "  Billable hrs  : $grandBill" -ForegroundColor Green
            Write-Host "─────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

            $allSummaries | Sort-Object TicketNumber | ForEach-Object {
               # $openMark = if (-not $_.Status -match 'complete') { ' *' } else { '' }
                Write-Host ("  {0,-20}  {1,-35}  {2,5}h  {3}" -f `
                        $_.TicketNumber, ($_.Title -replace '(.{33}).+', '$1…'), $_.TotalHours, $_.Status) `
                    -ForegroundColor Gray
            }
            Write-Host ""
        }

        # ── 5. HTML output ───────────────────────────────────────────────────────
        if ($OutputHtml -or $OpenInBrowser) {
            $htmlContent = Build-ATTicketHtml -Reports $reports -Dark:$DarkMode
            if (-not $HtmlPath) {
                $stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
                $HtmlPath = Join-Path $env:TEMP "ATTicketReport_$stamp.html"
            }
            $htmlContent | Out-File -FilePath $HtmlPath -Encoding UTF8 -Force
            Write-Host "Get-ATTicketReport: HTML saved to $HtmlPath" -ForegroundColor Cyan
            if ($OpenInBrowser) { Start-Process $HtmlPath }
        }

        # Return the array of report objects (one per ticket)
        return $reports.ToArray()
    }

    function Get-ATWeekStart {
        <#
    .SYNOPSIS
    Returns the date of the Sunday that was N weeks ago, at midnight.

    .DESCRIPTION
    Calculates the date of the Sunday at the start of the week that was LastXWeeks weeks ago.
    Weeks are considered to start on Sunday.  The returned DateTime has its time component
    set to 00:00:00 (midnight).

    This is used as the start-of-period date when querying time entries or calendar events
    for a rolling number of weeks.

    .PARAMETER LastXWeeks
    How many weeks back to calculate. Default is 1 (the most recent past Sunday).

    .EXAMPLE
    Get-ATWeekStart -LastXWeeks 1
    # Returns last Sunday at midnight

    .EXAMPLE
    Get-ATWeekStart -LastXWeeks 4
    # Returns the Sunday from 4 weeks ago at midnight

    .NOTES
    Sunday is treated as the last day of the week (value 7) for calculation purposes.
    Accepts pipeline input.
    #>
        [CmdletBinding()]
        param (
            [Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true, Mandatory = $false, Position = 0)]
            [int]$LastXWeeks = 0
        )
        if (!$LastXWeeks ) {
            $LastXWeeks = 0
        }
        #write-host "Get-ATWeekStart : this could be wrong if the current day is Sunday - need to check this" -ForegroundColor Yellow
        $wday = [int](Get-Date).DayOfWeek
    
        if ($wday -eq 0) { $wday = 7 } # convert Sunday from 0 to 7 for easier calculations
        #    (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-1 - ((7 * ($lastXWeeks - 1)) - $wday))
        write-verbose "Get-ATWeekStart: Today is $wday) which is weekday number $wday (with Monday as 1  and Sunday as 7). Calculating the Sunday of $LastXWeeks weeks ago." #-ForegroundColor Green
#        $i = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays( - ((7 * ($lastXWeeks - 1)) + $wday))
        $i = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays( - ((7 * ($lastXWeeks )) + $wday))
        Write-Verbose "Get-ATWeekStart: The Sunday of $LastXWeeks full weeks ago was on $($i.ToShortDateString())" #-ForegroundColor Green
        return $i




    }

    function Invoke-ATQuery {
        <#
    .SYNOPSIS
    Executes a paginated GET query against the Autotask REST API and returns all matching records.

    .DESCRIPTION
    Core read-only query engine for the AutoTaskRest module. Handles three calling modes:

      1. Raw URL  (-url)             — pass a fully-formed URL; returns the raw response body.
      2. Fixed suffix (-UrlFixedSuffix) — appends a path suffix to the saved base URL.
      3. Entity query (-entityName)  — builds a structured Autotask filter query and handles
         pagination automatically by recursing up to LoopCount times (default 40),
         collecting all pages of 500 records into a single result set.

    Authentication credentials are read from the kiss-atapi login file (written by Set-ATLogin).
    The secret is decrypted at runtime via DPAPI and is never written to any log.

    NOTE: Because PowerShell DateTime handling can cause locale-dependent formatting issues,
    use Convert-ObjArrayDateTimesToSearchableStrings on any result set that contains date fields
    before exporting to CSV or passing dates into filter expressions.

    .PARAMETER url
    A fully-formed Autotask API URL. Required for the 'raw' parameter set.
    Also used internally during recursive pagination calls for next-page URLs.

    .PARAMETER UrlFixedSuffix
    A path suffix appended directly to the saved base URL.
    Example: 'v1.0/Companies/29762985/Alerts'
    Required for the 'suffix' parameter set.

    .PARAMETER entityName
    The Autotask entity path to query. Example: 'v1.0/Companies', 'v1.0/TimeEntries'.
    Required for the 'entity' parameter set.

    .PARAMETER ID
    When provided (and not -1), retrieves a single record by its numeric Autotask ID.
    The constructed URL becomes: <baseUrl><entityName>/<ID>

    .PARAMETER isActive
    Switch. When set, appends an isActive = true filter clause to the query.

    .PARAMETER SearchFirstBy
    Controls the leading filter clause prepended to every entity query:
      id        — filter where id >= 0 (returns all records ordered by ID). Default.
      isActive  — filter where isActive = true.
      Nothing   — no leading filter; supply all filter logic via -SearchFurtherBy.

    .PARAMETER SearchFurtherBy
    One or more raw Autotask JSON filter clause strings appended after the SearchFirstBy clause.
    Multiple clauses must be comma-separated; the API implicitly AND-s them.
    Example: '{"op":"eq","Field":"companyName","value":"Acme Corp"}'

    .PARAMETER includeFields
    String array of field names to include in the response. When omitted, all fields are returned.
    Example: "id", "companyName", "isActive"

    .PARAMETER CheckDuplicatesOf
    Field name to inspect for duplicates across accumulated result pages.
    If duplicates are detected the function writes a warning and returns early,
    preventing inflated counts caused by API pagination anomalies.

    .PARAMETER LoopCount
    Maximum number of recursive pagination iterations before giving up. Default is 40
    (covers up to 20,000 records at 500 records per page). Reduce to add a tighter safety cap.

    .PARAMETER returnRaw
    Switch. When combined with -url, returns the full Invoke-RestMethod response object
    rather than extracting the .items array.

    .PARAMETER LoginInfo
    Optional PSCustomObject containing pre-built login details (url, UserName, Secret, atapi).
    When omitted the saved kiss-atapi credentials file is used.
    Primarily used internally by Test-ATConnection.

    .PARAMETER alreadyCapturedData
    Internal parameter passed during recursive pagination calls. Contains the records already
    fetched from earlier pages so that duplicate detection can operate across the full result set.
    Do not supply this parameter when calling Invoke-ATQuery directly.

    .EXAMPLE
    # Return all companies (all pages collected automatically)
    Invoke-ATQuery -entityName 'v1.0/Companies' -SearchFirstBy id

    .EXAMPLE
    # Return a single company record by ID
    Invoke-ATQuery -entityName 'v1.0/Companies' -ID 29762985

    .EXAMPLE
    # Return classification icons with a reduced field set
    Invoke-ATQuery -entityName 'v1.0/ClassificationIcons' -includeFields "id", "name"

    .EXAMPLE
    # Search companies by name fragment
    Invoke-ATQuery -entityName 'v1.0/Companies' -SearchFirstBy Nothing `
        -SearchFurtherBy '{"op":"contains","Field":"companyName","value":"Acme"}'

    .EXAMPLE
    # Fetch a child resource using a fixed URL suffix
    Invoke-ATQuery -UrlFixedSuffix "v1.0/Companies/29762985/Alerts"

    .EXAMPLE
    # Discover the regional API endpoint for a user (no credentials needed)
    Invoke-RestMethod "http://webservices.autotask.net/atservicesrest/v1.0/zoneInformation?user=api@example.com"

    .NOTES
    - The Autotask REST API caps responses at 500 records per page. This function recurses
      automatically via the pageDetails.nextPageUrl field to collect every page.
    - Credential secrets (password and API integration code) are decrypted from DPAPI-protected
      strings and are NEVER logged, even at -Verbose or -Debug verbosity levels.
    - For write operations (POST, PUT, PATCH, DELETE) use Invoke-ATREST instead.
    - See also: Set-ATLogin, Test-ATConnection, Convert-ObjArrayDateTimesToSearchableStrings.
    #>
        [CmdletBinding(DefaultParameterSetName = 'raw')]
        param (

            [Parameter(ParameterSetName = 'raw', Mandatory = $true)]
            [string]
            $url,

            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $false)]
            [string]
            $urlStart, # ='https://webservices6.autotask.net/atservicesrest/', #v1.0/',

            [Parameter(ParameterSetName = 'suffix', Mandatory = $true)]
            [string]
            $UrlFixedSuffix,
            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $true)]
            [string]
            $entityName,

            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $false)]
            [nullable[Int]]$ID = -1,

            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $false)]
            [switch]
            $isActive = $false,

            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $false)]
            [ValidateSet("id", "isActive", "Nothing")]
            [string]
            $SearchFirstBy = "id",

            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $false)]
            [string]
            $SearchFurtherBy,

            # Parameter help description
            [Parameter(ParameterSetName = 'entity', Mandatory = $false)]
            [string[]]
            $includeFields,
            [string]
            $CheckDuplicatesOf = $null,
            # Parameter help description
            [Parameter(Mandatory = $false)]
            [Int32]
            $LoopCount = 40,

            # [Parameter(ParameterSetName = 'raw', Mandatory = $false)]
            [switch]
            $returnRaw = $false,

            [PSCustomObject]$LoginInfo,
            [PSCustomObject]$alreadyCapturedData


            # [string]$apiUsername,
            # [string]$apiPassword,
            # [string]$apiID

        )

        # $saveobj = @{
        #     atapi    = ''#ConvertFrom-SecureString -SecureString $l_Apiid
        #     UserName = ''#"$apiusername"
        #     Secret   = '' #ConvertFrom-SecureString -SecureString $l_secret
        #     url      = ''# "$($r.url)"
        # }
        if ($alreadyCapturedData) {
            if ($CheckDuplicatesOf) {
                write-verbose "i-ATAPI: checking for duplicate values"
                $arethereduplicates = $alreadyCapturedData | Group-Object $CheckDuplicatesOf
                if ($arethereduplicates.Count -ne $alreadyCapturedData.Count) {
                    #   if  ($arethereduplicates.Count -gt 1){
                    write-host "I-AutotaskAPI $($arethereduplicates.Count) duplicates exists"
                    write-host "I-AutotaskAPI did not return all values"
                    #throw "NOT ALL DATA RETURNED, $CheckDuplicatesOf has duplicates"
                    return

                }

            } 
        }

        # Build auth header via central helper — decrypts secret and atapi, zeros BSTR buffers immediately
        $baseUrl = $null
        try {
            $kissATheader = Get-ATCredentialHeader -LoginInfo $LoginInfo -BaseUrl ([ref]$baseUrl)
        }
        catch {
            Write-Warning "Invoke-ATQuery: $_"
            throw
        }

        if ($url -and ($returnRaw -eq $true)) {
            Write-Verbose "Invoke-ATQuery get RAw data based on $url"
            Invoke-RestMethod -Method Get -Uri $url  -Headers $kissATheader  #-SkipHeaderValidation
            Write-Debug "url: $url"  # header intentionally excluded from debug output to avoid logging credentials
            return
        }
        if ($urlFixedSuffix) {
            $url2 = "$baseUrl$UrlFixedSuffix"
            Write-Verbose "Invoke-ATQuery get Raw data based on $url2"
            Invoke-RestMethod -Method Get -Uri $url2  -Headers $kissATheader    #-SkipHeaderValidation 
            return


        }
   
        if (($id -ne -1) -and ($null -ne $id) -and $entityName) { 
            # just return a SINGLE item with a specific ID
            # $url2 = "$urlstart$entityName/$ID"
            $url2 = "$baseUrl$entityName/$ID"
            Write-Verbose "Invoke-ATQuery getiing just one $entityname item $id : $url2"
            $Result = Invoke-RestMethod -Method Get -Uri $url2  -Headers $kissATheader  #-SkipHeaderValidation #-FollowRelLink
            Write-Debug "url: $url2"  # header intentionally excluded from debug output to avoid logging credentials

            Write-Verbose "Invoke-ATQuery item count=$($result.item.count)"
            if ($ReturnRaw -eq $true) {
                write-host "Invoke-ATQuery Returning raw data, and not an object collection - this WILL include userDefinedFields"
                return $result
            }
            return $Result.Item
        }
 
        if ($entityName) {
            # prepare a collection of items to return - and might need to be called recursively
            $entityFilter = ''
            switch ($SearchFirstBy) {
                "isActive" {
                    Write-Verbose "Invoke-ATQuery : returning only $entityname items where field isActive = true"
                    $entityfilter = '{"op":"eq","field":"isActive","value":"true"}'
                }
                "id" {
                    Write-Verbose "Invoke-ATQuery : returning  $entityname where ID GTE 0 and isactive:$isactive"
                    $entityfilter = '{"op":"gte","field":"id","value":"0"}'
                    if ($isActive) {
                        $entityfilter += ',{"op":"eq","field":"isActive","value":"true"}'
                    }
              
                }
                Default {
                    if ($isActive) {
                        $entityfilter += '{"op":"eq","field":"isActive","value":"true"}'
                    }
                }
            }
 
            $entityfilter = "$entityfilter,$SearchFurtherBy".trim(',')
            $_search = """filter"":[$entityfilter]"
            if ($includeFields) {
                $includeFields = ('"{0}"' -f ($includeFields -join '","'))  # turn an array into a quoted comma seperated list
        
                $_search = """IncludeFields"":[$includefields],$_search"
            
            }
            $_search = $_search.replace('""', '"')
            $url2 = "$baseUrl$entityName/query?search={$_search}"
            #$url2 = "$urlstart$entityName/query?search={$_search}"
        }
        else { $url2 = $url }
    
        Write-verbose "getting  $entityname items  $url2"
        $Result = Invoke-RestMethod -Method Get -Uri $url2  -Headers $kissATheader  #-SkipHeaderValidation
        $RecordsRecieved = $Result.pageDetails.Count
        $apidata = $Result.items
        $apidata
        Write-Verbose "retrieved $RecordsRecieved records: which should equal $($apidata.count)"
        Write-Verbose "returned PageDetails `n$($Result.pageDetails |ConvertTo-Json)"

    
        #now prepare the next 500 items
        $Nextpage = $Result.pageDetails.nextPageUrl
        if (($LoopCount -gt 1) -and $Nextpage) {
            Write-Verbose "Invoke-ATQuery LoopCount Value = $Loopcount"

            if ($CheckDuplicatesOf) {
                $alreadyCapturedData += $apidata
                $apidataT = Invoke-ATQuery -url $Nextpage -LoopCount ($LoopCount - 1) -CheckDuplicatesOf $CheckDuplicatesOf -alreadyCapturedData $alreadyCapturedData
                $apidata += $apidataT
                $apidataT
            }
            else {
                $apidataT = Invoke-ATQuery -url $Nextpage -LoopCount ($LoopCount - 1)
                $apidata += $apidataT
                $apidataT
            }


        }
        Write-Verbose "Invoke-ATQuery total Returned items = $($apidata.count)"
        return 
    }


    function Invoke-ATREST() {
        <#
    .SYNOPSIS
    Low-level REST wrapper for Autotask API calls that write or delete data (PUT, POST, PATCH, DELETE).

    .DESCRIPTION
    Reads saved credentials from the kiss-atapi login file, builds the required Autotask
    authentication header, then calls Invoke-RestMethod with the supplied HTTP method, URL
    suffix, and optional JSON body.  Use this function for any operation that mutates data
    (create, update, patch, delete).  For read-only queries prefer Invoke-ATQuery, which
    handles pagination automatically.

    Credentials must have been saved previously with Set-ATLogin.

    .PARAMETER url
    The URL path appended to the base Autotask URL stored in the login file.
    Example: 'V1.0/Companies/12345/Alerts'

    .PARAMETER Body
    Optional JSON body string for POST, PUT, and PATCH requests.
    Obtain this by converting a PSCustomObject: $obj | ConvertTo-Json -Compress

    .PARAMETER Method
    The HTTP verb to use.  Must be one of: PUT, GET, POST, DELETE, PATCH.

    .EXAMPLE
    # Patch a contact's email address
    $json = [PSCustomObject]@{ id = 12345; emailAddress = "new@example.com" } | ConvertTo-Json -Compress
    Invoke-ATREST -url 'V1.0/Companies/99/Contacts' -Method PATCH -Body $json

    .EXAMPLE
    # Delete a company alert
    Invoke-ATREST -url 'V1.0/Companies/99/Alerts/77' -Method DELETE

    .NOTES
    Credentials are read from $kissATAPIpath\$kissATAPIfile on every call.
    The secret is stored as an encrypted string and is decrypted at runtime via DPAPI.
    #>
        [CmdletBinding()]
        param (
            [Parameter(ParameterSetName = 'raw', Mandatory = $true)]
            [string]
            $url,
            [Parameter( Mandatory = $false)]
            [string]
            $Body,
            [Parameter( Mandatory = $true )]
            [ValidateSet("PUT", "GET", "POST", "DELETE", "PATCH")]
            [string]
            $Method
        )
        $baseUrl = $null
        $kissATheader = Get-ATCredentialHeader -BaseUrl ([ref]$baseUrl)
        write-debug "Invoke-ATREST: base url is $baseUrl"
        $url2 = "$baseUrl$Url"
        write-verbose "Invoke-ATREST $Method $url2 `r`n BODY $body"
        $result = Invoke-RestMethod -Method $Method -Uri $url2 -Headers $kissATheader -Body $Body
        write-verbose "Invoke-ATREST resultitem = $($result.itemid)"
        $result

    }

    function New-AT365CalendarEvent {
        <#
    .SYNOPSIS
    Creates a new calendar event in the current user's Microsoft 365 calendar via Microsoft Graph.

    .DESCRIPTION
    Connects to Microsoft Graph with Calendars.ReadWrite scope and creates a new calendar event
    for the authenticated user.  Reminders are disabled by default.
    If AutotaskTimeEntryID is supplied (and is not -1), it is stored as a custom single-value
    extended property on the event using a fixed GUID, allowing Sync-AT365Calendar
    to match events back to Autotask time entries in future runs.

    StartUTC must be before EndUTC — the function exits with an error message if this is violated.

    .PARAMETER Subject
    The subject line of the calendar event.

    .PARAMETER Body
    Plain-text body content for the event. Typically contains the TimeEntry ID, ticket details,
    and hours worked so the event can be matched and updated later.

    .PARAMETER StartUTC
    The start date and time of the event, expressed in UTC.

    .PARAMETER EndUTC
    The end date and time of the event, expressed in UTC. Must be after StartUTC.

    .PARAMETER AutotaskTimeEntryID
    Optional. The numeric Autotask Time Entry ID to embed in the event's extended properties.
    Default is -1 (not set).

    .EXAMPLE
    New-AT365CalendarEvent -Subject "AT: Acme Corp - Fix server" -Body "TimeEntry:12345`nHours: 1.5" -StartUTC (Get-Date).ToUniversalTime() -EndUTC (Get-Date).AddHours(1.5).ToUniversalTime() -AutotaskTimeEntryID 12345

    .NOTES
    Requires the Microsoft.Graph.Calendar module and Calendars.ReadWrite scope.
    The extended property GUID is hard-coded as 1e388ea9-5c0d-4aec-aaf9-8150a6e7797c.
    #>
        [CmdletBinding()]
        param (
            [string]$Subject,
            [string]$Body,
            [DateTime]$StartUTC,
            [DateTime]$EndUTC,  
            [Int64]$AutotaskTimeEntryID = -1

        )
    
        if (-not(Get-Module -ListAvailable -Name  Microsoft.Graph)) { 
            # if (-not(Get-InstalledModule Microsoft.Graph)) { 
            #Get-Module -ListAvailable -Name Microsoft.Graph  
            Write-Host "Microsoft Graph module not found" -ForegroundColor Black -BackgroundColor Yellow
            $install = Read-Host "Do you want to install the Microsoft Graph Module?"
  
            if ($install -match "[yY]") {
                Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
            }
            else {
                Write-Host "Microsoft Graph module is required." -ForegroundColor Black -BackgroundColor Yellow
                throw "Microsoft Graph module is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
            } 
        }
        Connect-MgGraph -Scopes  "Calendars.ReadWrite" | Out-Null

        $365me = (Get-MgContext).Account
        write-verbose "New-AT365CalendarEvent: connected to Microsoft Graph as $365me, now creating a new calendar event with subject '$Subject'"
        if ($startUTC -ge $endUTC) {
            write-host "New-AT365CalendarEvent: Error - StartUTC must be before EndUTC" -ForegroundColor Red
            return
        }

        # my random guid 1e388ea9-5c0d-4aec-aaf9-8150a6e7797c
        #Used by 365 MgGraph tool
        #$pa = $item.PropertyAccessor
        #$propTag = "http://schemas.microsoft.com/mapi/string/{1e388ea9-5c0d-4aec-aaf9-8150a6e7797c}/AutotaskTimeEntryID"
        #$item.PropertyAccessor.SetProperty($propTag, "$($timeEntry.id)")

        New-MgUserEvent -UserId $365me    -BodyParameter @{
            subject                       = $subject
            start                         = @{
                dateTime = $startUTC.ToString("o") # ISO 8601 format for date-time
                timeZone = "UTC"
            }
            end                           = @{
                dateTime = $EndUTC.ToString("o") # ISO 8601 format for date-time
                timeZone = "UTC"
            }
            isReminderOn                  = $false
            Body                          = @{
                contentType = "Text"
                content     = $Body
            }
            singleValueExtendedProperties = @(
                @{
                    id    = "String {1e388ea9-5c0d-4aec-aaf9-8150a6e7797c} Name AutotaskTimeEntryID"
                    value = $AutotaskTimeEntryID
                }
            )


        } | Out-Null


        write-verbose "New-AT365CalendarEvent: created new calendar event with ID $($newEvent.Id) Subject: $($newEvent.Subject) Start: $($newEvent.Start.DateTime) End: $($newEvent.End.DateTime)"
        return 
    }

    function Set-ATCompanies() {
        <#
    .SYNOPSIS
    Updates Autotask company records with branch, manager, and classification information.

    .DESCRIPTION
    Updates one or more Autotask companies with Primary/Secondary engineer assignments,
    branch location, manager, and classification.
    - If a field parameter is "", no action is taken for that field.
    - If the secondary field is "" or "null", any existing secondary assignment is removed.
    - Companies must be identified by either CompanyID or an EXACT CompanyName match.
    Accepts pipeline input from a CSV import for bulk updates.

    .PARAMETER CompanyID
    The unique numeric Autotask company ID. Alias: ID.

    .PARAMETER CompanyName
    The exact company name as it appears in Autotask. Used as an alternative to CompanyID.
    Aliases: Name, Company.

    .PARAMETER Manager
    The manager to assign to the company. Can be a name string or a numeric resource ID.

    .PARAMETER Classification
    The classification to apply. Can be a name string (e.g. 'Residential') or a numeric ID.

    .PARAMETER Branch
    The branch to assign. Can be a branch name string or a numeric branch/location ID.

    .EXAMPLE
    Set-ATCompanies -CompanyID 0 -branch "Tauranga - Kiss I.T" -Verbose

    .EXAMPLE
    Set-ATCompanies -CompanyName "Kiss IT" -branch 29682914 -Manager "Sean Macey" -Classification Residential

    .EXAMPLE
    $CSV = Import-Csv .\companies.csv
    $CSV | Set-ATCompanies

    .NOTES
    Accepts pipeline input with ValueFromPipelineByPropertyName, so CSV columns named
    CompanyID, CompanyName, Manager, Classification, and Branch are automatically mapped.
    #>
        [CmdletBinding()]
        param (
       
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [int[]]
            [alias("ID")]
            $CompanyID = -1,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            [alias("Name")]
            [alias("Company")]
            $CompanyName,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            $Manager,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            $Classification ,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            $Branch,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [bool]
            $isActive



        )
        begin {
            if ($Classification -and ($Classification -ne "null")) { $classes = Get-ATClassificationIcons }
            if ($Manager -and ($Manager -ne "null")) { $Engineers = Get-ATEngineers -isActive }
            if ($Branch -and ($Branch -ne "null")) { $Branches = ( Invoke-ATREST -Method GET -url '/V1.0/UserDefinedFieldListItems/query?search={"filter":[{"op":"eq","field":"udfFieldId","value":"29682941"}]}' ).items }
            #  $ipatch = 0
            # $patchObj =@()
    
            <#
      UserDefinedFieldDefinitions  . Branch => id = 29682941 (datatype 3)
      #>
        }
        process {
 

            write-verbose "Set-ATCompanies CompanyID to process = $CompanyID"
            $checkids = $CompanyID
            # if ($Manager -eq "null"){ $Manager = ""}
            # if ($CompanyType -eq "null") { $CompanyType = "" }
            if (($checkids -eq -1) -and $CompanyName) {
            
              #  $escapedName = convertto-escapedString -inputString $CompanyName
           $escapedName =  [Uri]::escapeDataString( $companyName )

              Write-Verbose "Set-ATCompanies About to check by name of Company $CompanyName : Comnpany ID= $checkids"
                $res = Get-ATCompanies -CompanyName $escapedName -exactNameMatch -DontExpandChildIDFields
                if ($Res) { $checkids = $res.ID }

            }
        
            If ($checkids -eq -1) { return }      
            foreach ($anID in $checkids) {
                $obj = [PSCustomObject]@{
                    id = -1
                }

                if ($Manager -gt 0) {
                    #  write-host "checking manager $Manager"
                    if ($Engineers.id -contains $Manager ) {
                        $obj.id = $anID
                        write-verbose " changing manager by ID = $Manager"
                        $obj | Add-Member -NotePropertyName "ownerResourceID" -NotePropertyValue $Manager
                    }
                    elseif ($Manager -eq "null") {
                        $obj.id = $anID
                        #$obj.Classification = ""}
                        write-verbose " Changing manager by NULL"
                        $obj | Add-Member -NotePropertyName "ownerResourceID" -NotePropertyValue ""
                    }
                    else {
                        $val = ""
                        $res = $Engineers | Where-Object FullName -eq $Manager
                        if ($Res) {
                            $val = $res.id
                            $obj.id = $anID
                            write-verbose " Changing manager by Fullname $Manger = ID $val "
                            $obj | Add-Member -NotePropertyName "ownerResourceID" -NotePropertyValue $val
                        }
                        else {
                            throw "Set-ATCompanies: Can not fully update CompanyID $anID : could not find Engineer/Manager in autotask matching $Manager "
                        }                      
                       
               
                    }
                }
                if ($Classification) {
                    if ($classes.id -contains $Classification) {
                        $obj.id = $anID
                        $obj | Add-Member -NotePropertyName "classification" -NotePropertyValue $Classification
                    }
                    elseif ($Classification -eq "null") {
                        $obj.id = $anID
                        $obj | Add-Member -NotePropertyName "classification" -NotePropertyValue ""
                    }
                    else {
                        $res = $classes | Where-Object name -like $Classification
                        if ($res) {
                            $val = $res.id
                            $obj.id = $anID
                            $obj | Add-Member -NotePropertyName "classification" -NotePropertyValue $val
                        }
                        else {
                            throw "Set-ATCompanies: Cannot update CompanyID $anID — no classification found in Autotask matching '$Classification'"
                        }
                    }
                

           
                }
                if ($Branch) {
                    # if ($Branches.id -contains $Branch) {$Branch = ($Branches |Where-Object id -eq $Branch).valueFor}
                    $val2 = $Branch
                    if ($Branches.id -contains $Branch) {
                        $val2 = ($Branches | Where-Object id -eq $Branch)[0].valueforExport
                    }
                    if ($Branches.valueforDisplay -contains $Branch) {
                        $val2 = ($Branches | Where-Object valueforDisplay -eq $Branch)[0].valueforExport
                    }


                    if (($Branches.valueforExport -contains $val2) -or ($Branches.valueforDisplay -contains $val2)) {
                        $obj.id = $anID
                        $userDefinedFields = @()
                        $v = [PSCustomObject]@{
                            Name  = "Branch"
                            value = $val2
                        }
                        if ($Branch -eq "null") { $v.value = "" }
                        $userdefinedFields += $v
                        $obj | Add-Member -NotePropertyName userDefinedFields -NotePropertyValue $userdefinedFields
                    }


                }
                if ($null -ne $isActive) {
                    $obj.id = $anID
                    $obj | Add-Member -NotePropertyName isActive -NotePropertyValue $isActive
                
                }
                if ($obj.id -ge -1) {
                    $json = $obj | ConvertTo-Json -Compress
                    write-Host "Set-AutotaskCompany update  $obj"
                    Invoke-ATREST -url 'V1.0/Companies' -Method PATCH -Body $json | Out-Null
              
                    # $patchObj += $Obj
                    # if ($patchObj.count -gt 200){
                    #     $json = ($patchObj | ConvertTo-Json -Compress).trim("[").trim("]")
                    #     Write-verbose " Set-AutotaskCompany update Json body $json"
                    #     Invoke-ATREST -url 'V1.0/Companies' -Method PATCH -Body $json | Out-Null
                    #     $patchObj = @()
                    # }

                }

            
            }
        }
    
        end {
            # if ($patchObj.count -gt 0){
            #     $json = ($patchObj | ConvertTo-Json -Compress).trim("[").trim("]")
            #     Write-verbose " Set-AutotaskCompany update Json body $json"
            #     Invoke-ATREST -url 'V1.0/Companies' -Method PATCH -Body $json | Out-Null
            # }
        }



    }

    function Set-ATCompanyEngineers() {
        <#
    .SYNOPSIS
    Updates the primary and secondary engineer assignments for one or more Autotask companies.

    .DESCRIPTION
    Sets or clears the Primary and Secondary engineer alert text stored against a company in Autotask.
    - If Primary is "" or "null", any existing primary assignment is removed.
    - If Secondary is "" or "null", any existing secondary assignment is removed.
    - Companies must be identified by CompanyID or an EXACT CompanyName match.
    - Accepts pipeline input so a CSV with columns CompanyID/Name, Primary, Secondary can be piped directly.

    Internally reads and writes alert records of type 1, 2, and 3 on the company, preserving
    any non-engineer text already present in those alerts.

    .PARAMETER CompanyID
    The unique numeric Autotask company ID. Alias: ID.

    .PARAMETER CompanyName
    The exact company name in Autotask. Used as an alternative to CompanyID.
    Aliases: Name, Company.

    .PARAMETER Primary
    The name of the primary engineer to assign. Pass "" or "null" to remove the assignment.

    .PARAMETER Secondary
    The name of the secondary engineer to assign. Pass "" or "null" to remove the assignment.

    .EXAMPLE
    $eng = Import-Csv .\PrimaryEngineers.csv
    $eng | Set-ATCompanyEngineers

    .EXAMPLE
    Set-ATCompanyEngineers -CompanyName "Matamata Medical Center" -Primary "Sean" -Secondary "Antony"

    .EXAMPLE
    Set-ATCompanyEngineers -ID 29762990 -Primary "Sean" -Secondary "null"

    .NOTES
    Uses Get-ATCompanyChildAlerts, and then PUT/POST/DELETE via Invoke-ATREST.
    The engineer name is stored as plain text in the alert body — it is not validated against
    the Autotask resources list.
    #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [int[]]#[int[]]
            [alias("ID")]
            $CompanyID = -1,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            [alias("Name")]
            [alias("Company")]
            $CompanyName,
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            $Primary = "",
            [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
            [string]
            $Secondary = ""


        )
        begin {
            #$i = 0
            #$jsontxt =""
        }
        process {
            $CheckIDs = $CompanyID
            if (($CheckIDs -eq -1) -and $CompanyName) {
#                $escapedName = convertto-escapedString -inputString $CompanyName 
                $escapedName = [Uri]::escapeDataString($CompanyName )


                $res = Get-ATCompanies -CompanyName $escapedName -exactNameMatch -DontExpandChildIDFields
                if ($res) { $CheckIDs = $res[0].id }
            }
            If ($CheckIDs -eq -1) {
                Write-Verbose "Set-ATCompanyEngineers: No CompanyID or CompanyName provided to identify the company to update - exiting"
                return 
            }

            if ($primary -eq "null") { $primary = "" }
            if ($secondary -eq "null") { $Secondary = "" }

            write-verbose "Set-ATCompanyEngineers CompanyID to process = $CheckIDs and CompanyName = $CompanyName "
            foreach ($anID in $CheckIDs) {
                write-host "modify PrimaryEngineers of $anID $CompanyName "
                $ChildAlerts = Get-ATCompanyChildAlerts -CompanyID $anID
                $x = 1
                $a = @(1, 2, 3)
                foreach ($x in $a) {
                    $alert = $ChildAlerts | Where-Object alertTypeID -eq $x



                    if ($alert) {
                        #Write-Verbose "write-CompanyPrimary alertTypeID:$x updating an existing alert record"

                        #must PUT
                        $json = $alert | ConvertTo-Json    
                        Write-Verbose  "write-CompanyPrimary alertTypeID:$x  initial data exists company $anID and alertType $x and =  $json"
                        $assignedTech = [PSCustomObject]@{
                            CompanyID      = $anID
                            Primary        = $null
                            Secondary      = $null
                            TextPrimary    = ""
                            TextSecondary  = ""
                            CompanyAlertID = $null
                        
                        }
                        if ($alert.AlertText -imatch "secondary\s+tech.*[:][\s|\w]*\n|secondary\s+engineer.*[:][\s|\w]*\n|secondary\s+tech.*[:][\s|\w]*|secondary\s+engineer.*[:][\s|\w]*") {
                            $assignedTech.TextSecondary = ($Matches[0]) -replace ("\n", "")
                            $assignedTech.CompanyAlertID = $l.ID
                            $assignedTech.secondary = $assignedTech.Textsecondary -ireplace [regex]::Escape("secondary"), ""
                            $assignedTech.secondary = $assignedTech.secondary -ireplace [regex]::Escape("engineer"), ""
                            $assignedTech.secondary = $assignedTech.secondary -ireplace [regex]::Escape("tech"), ""
                            $assignedTech.secondary = $assignedTech.secondary.replace(":", "").trim()
                        } 

                        if ($alert.AlertText -imatch "primary\s+tech.*[:][\s|\w]*\n|primary\s+engineer.*[:][\s|\w]*\n|primary\s+tech.*[:][\s|\w]*|primary\s+engineer.*[:][\s|\w]*") {
                            $assignedTech.TextPrimary = ($Matches[0]) -replace ("\n", "") 
                            $assignedTech.CompanyAlertID = $anID
                            $assignedTech.Primary = $assignedTech.TextPrimary -ireplace [regex]::Escape("primary"), ""
                            $assignedTech.Primary = $assignedTech.Primary -ireplace [regex]::Escape("engineer"), ""
                            $assignedTech.Primary = $assignedTech.Primary -ireplace [regex]::Escape("tech"), ""
                            $assignedTech.Primary = $assignedTech.Primary.replace(":", "").trim()

                        }
                        $atemp = $alert.alertText -replace ($assignedTech.TextPrimary, "") -replace ($assignedTech.TextSecondary, "").Trim() -replace '^(\n)*', ""
                        #if ($atemp) {$atemp = $atemp -replace '^(\n)*',""}
                        $alert.alertText = ""
                        if ($primary -and ($x -ne 2)) {
                          #  $escapedPrimary = convertto-escapedString -inputString $primary
                           $escapedPrimary = [Uri]::escapeDataString( $primary )
                            $alert.alertText = "Primary Engineer: $escapedPrimary`n"
                        }

                        if ($secondary -and ($x -ne 2)) {
                            $escapedSecondary = [Uri]::escapeDataString( $secondary )
                            $alert.alertText = $alert.alertText + "Secondary Engineer: $escapedSecondary"
                        }


                        if ($atemp) {
                            Write-Verbose "write-companyPrimary: alerttypeid:$x found extra text $atemp"
                            $alert.alertText = $alert.alertText.trim() + "`n" + "$atemp"
                        }
                        $alert.alertText = $alert.alertText -replace '^(\n)*', "" #-replace "^(`n",""
                        if ($alert.alertText) {
                            #the alert exists - so update it
                            $json = $alert | ConvertTo-Json 
                            write-verbose "write-CompanyPrimary alertTypeID:$x Updating Primary and secondary Engineer for $anID"
                            Invoke-ATREST -url ('V1.0/Companies/' + $anID + '/Alerts') -Method PUT -Body $json | Out-Null
                            # $jsontxt += $json
                        }
                        elseif ($alert.id) {
                            #there is no needed alertText, so DELETE the alert
                            write-verbose "write-CompanyPrimary alertTypeID:$x Deleting Primary Engineer for $anID"
                            Invoke-ATREST -url ('V1.0/Companies/' + $anID + '/Alerts/' + $alert.id) -Method DELETE  | Out-Null
                            #  $jsontxt += $json

                        }

                    }
                    else {
                        if (($Primary -or $Secondary) -and ($x -ne 2)) {
                            #creating a new alert
                            Write-Verbose "write-CompanyPrimary alertTypeID:$x creating a NEW Primary/Secondary Engineer record"
                            $alert = [PSCustomObject]@{
                              #  alertText   = convertto-escapedString -inputString "Primary Engineer: $primary`nSecondary Engineer:$secondary"
                                alertText   = [Uri]::escapeDataString("Primary Engineer: $primary`nSecondary Engineer:$secondary")
                                alertTypeID = $x
                                companyID   = $anID
                            } 
                            $json = $alert | ConvertTo-Json
                            Invoke-ATREST -url ('V1.0/Companies/' + $anID + '/Alerts') -Method POST -Body $json | Out-Null
                            #  $jsontxt += $json

                        }
                    }
                }
                # $i = $i + 1
                # if ($i -gt 10)
                # {
                #     $i = 0
                #     write-Host "Set primary - expect loop \n $jsontxt"
                # }
            }
        
        }

        end {
            #write-host "set primary: Now finsih everything \n $jsontxt"
        }


        # $json = $alert | ConvertTo-Json    
        # Write-Host "$json"
        # $assignedTech
    }

    function Set-ATContact {
        <#
    .SYNOPSIS
    modify a contact
    
    .DESCRIPTION
    update contact infromations (email and whether contact has opted pout of bulk emails)
    will accept an array of contact objects to be piped into it (so can process bulk contacts from CSV)
    
    .PARAMETER Contact
    Parameter description
    
    .PARAMETER eMail
   email address
    
    .PARAMETER SetunknownEmail
    if this is true, then set the email to unknown@unknown.co.nz
    
    .PARAMETER isOptedOutFromBulkEmail
    sets (or unsets) the contact from bulkemail outs
    
    .EXAMPLE
    $a = Get-ATContacts -eMail trevor@belvedereconstruction.co.nz
    $a | Set-ATContact -isOptedOutFromBulkEmail False
    $a | Set-ATContact -isOptedOutFromBulkEmail $true
    
    .NOTES
    General notes
    #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [psobject]
            $Contact,
            [Parameter(Mandatory = $false)]
            [string]
            $eMail,
            [Parameter(Mandatory = $false)]
            [switch]
            $SetunknownEmail,
            [Parameter(Mandatory = $false)]
            [string]
            [validateset("True", "False", "NoChange")]
            $isOptedOutFromBulkEmail


        )

        begin {
        
        }
    
        process {
            if (($Contact.id -gt -1) -and ($contact.companyID -gt -1)) {
                $companyID = $contact.companyID
                $obj = [PSCustomObject]@{
                    id = $Contact.id
                    #  emailAddress = "unknown@unknown.co.nz"
                } 
                if ($SetunknownEmail -eq $true) {
                    $obj | Add-Member -NotePropertyName "emailAddress" -NotePropertyValue "unknown@unknown.co.nz"
                    # $obj = [PSCustomObject]@{
                    #     id           = $Contact.id
                    #     emailAddress = "unknown@unknown.co.nz"

                    # } 
                    if (!$contact.emailAddress3) {
                        $obj | Add-Member -NotePropertyName "emailAddress3" -NotePropertyValue $Contact.emailAddress
                    }
                }

                if ($isOptedOutFromBulkEmail -eq "True") {
                    write-verbose "set contact opted out TRUE"
                    $obj | Add-Member -NotePropertyName "isOptedOutFromBulkEmail" -NotePropertyValue "True"
                }
                if ($isOptedOutFromBulkEmail -eq "False") {
                    write-verbose "set contact opted out FALSE"
                    $obj | Add-Member -NotePropertyName "isOptedOutFromBulkEmail" -NotePropertyValue "FALSE"
                }
            
                $json = $obj | ConvertTo-Json -Compress 
                write-Host "Set-ATContact  $obj"
                Invoke-ATREST -url "V1.0/Companies/$companyID/Contacts" -Method PATCH -Body $json  | Out-Null
    
            }
        }
        end {
    
        }

    }



    function Set-ATEngineerToReport() {
        <#
    .SYNOPSIS
    Saves an engineer's Autotask Resource ID to the local login file so it can be used as the default for reporting.

    .DESCRIPTION
    Looks up an active engineer in Autotask by either their numeric Resource ID or their email address,
    then writes the found ATResourceID value into the kiss-atapi login JSON file.
    Once saved, Get-ATEngineerToReport returns this value and functions such as
    Get-ATTimeEntries -ForMeOnly use it to filter results to that engineer.

    .PARAMETER id
    The numeric Autotask Resource ID of the engineer. Use this parameter set (ByID) when you already
    know the engineer's ID.

    .PARAMETER email
    The email address of the engineer in Autotask. Use this parameter set (ByEmail) to look up the
    engineer without knowing their ID first.

    .EXAMPLE
    Set-ATEngineerToReport -id 29683001

    .EXAMPLE
    Set-ATEngineerToReport -email "jane.smith@kissit.co.nz"

    .NOTES
    Requires Set-ATLogin to have been run first so the login file exists.
    Only active engineers are accepted — if the ID or email does not match an active resource,
    nothing is saved and a warning is displayed.
    #>
        [CmdletBinding(DefaultParameterSetName = "ByID")]
        param (
            [Parameter(ParameterSetName = 'ByID', ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory = $true)]
            [int]
            $id,
            [Parameter(ParameterSetName = 'ByEmail', ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory = $true)]
            [string]
            $email
        )
        Process {
            $engineer = $null
            if ($PSBoundParameters.ContainsKey('id')) {
                Write-Host "Set-AutoTaskEngineerToReport: Searching for Engineer with ID $EngineerID"
                $engineer = Get-ATEngineers -id $id -isActive
            }
            elseif ($PSBoundParameters.ContainsKey('email')) {
                Write-Host "Set-AutoTaskEngineerToReport: Searching for Engineer with email $email"
                $engineer = Get-ATEngineers -email $email -isActive
            }
            else {
                Write-Host "Set-AutoTaskEngineerToReport: No valid parameter provided. Please provide either EngineerID or email." -ForegroundColor Red
                return
            }

            if ($engineer) {
                Write-Host "Set-AutoTaskEngineerToReport: Found Engineer: $($engineer.FullName) with ID $($engineer.ID)"
         
                # Here you can add code to set the engineer to report, e.g., update a database, send an email, etc.
                # if (!(Test-Path -Path $kissATAPIpath)) {
                #     new-item -Path $home -Name kiss-atapi -ItemType Directory
                #     Write-Host "Created a new Directory called $($home)\kiss-atapi" 
                # }
                # else {
                if (test-path -path "$kissATAPIpath\$kissATAPIfile" ) {
                    $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
                    if ($jsn) {
                        #  write-host "there was a prexisting saved login of $jsn"
                        $r = $jsn  | ConvertFrom-Json 
                        if ($r | Get-Member -Name "ATResourceID" -ErrorAction SilentlyContinue) {
                            # if ($r.PSObject.Properties.Match("ATResourceID")) {


                            write-Host "Updating the saved login with the new ATResourceID of $($engineer.ID)" -ForegroundColor Green
                            $r.ATResourceID = $engineer.ID
                        }

                        else {
                            # Create it if missing
                            write-Host "No ATResourceID found in the saved login, adding this property with the new ATResourceID of $($engineer.ID)" -ForegroundColor Yellow
                            $r | Add-Member -MemberType NoteProperty -Name "ATResourceID" -Value $engineer.ID
                        }
                        
                        $r | ConvertTo-Json -Compress | Out-File "$kissATAPIpath\$kissATAPIfile" -Force
                        Write-Host "Updated the saved login with the new ATResourceID of $($engineer.ID)"
                    }

                }
                # }


            }
            else {
                Write-Host "Set-AutoTaskEngineerToReport: No active engineer found with the provided identifier." -ForegroundColor Yellow
            }
        }
    

    }
    function Set-ATInternalTicketTime() {
        <#
    .SYNOPSIS
    Annotates time entries that belong to internal company tickets with internal-hours classification fields.

    .DESCRIPTION
    Examines the supplied time-entry array and identifies entries whose TicketID matches a ticket
    belonging to a known set of internal Kiss IT companies.  For each matching entry the function
    populates four extra properties:
      - InternalTicketBillableNormalHrs
      - InternalTicketNonBillableNormalHrs
      - InternalTicketBillableAftHrs
      - InternalTicketNonBillableAftHrs
    and a combined InternalTicket total.

    The input object is modified in place (it is passed by reference), so the function does not
    need to be captured in a variable — changes appear in the original array.

    .PARAMETER timeEntries
    A PSObject array of time entries, as returned by Get-ATTimeEntries.
    Must contain at minimum the properties: TicketID, CompanyID, dateWorked.

    .EXAMPLE
    $entries = Get-ATTimeEntries -LastxMonths 1
    Set-ATInternalTicketTime -timeEntries $entries
    # $entries now contains the internal-ticket hour columns

    .NOTES
    The list of internal company IDs is hard-coded: 29762985, 0, 1, 29740186, 29761818, 29762138,
    29718567, 29762986.  Update these values if the set of internal companies changes.
    The $ATnonBillableCodes variable is expected to be available in the calling scope.
    #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [psobject]
            $timeEntries
            #must fields : ticketID, CompanyID, dateworked
        
        )
        # this function modifies the $timeEntries obj, it does not need to return it! since the input object is reference, not a value
        if (!($timeEntries.TicketID -and $timeEntries.CompanyID -and $timeEntries.dateWorked)) {
            throw "the timeentries input object is missing either TicketID, CompanyID or dateWorked"
        }

        #insert the username of eachtech
        $Resources = Get-ATEngineers #| Where-Object { ($_.id -in $TimeEntries.resourceID) }
        $timeEntries | Add-Member -NotePropertyName 'Resource' -NotePropertyValue "unknown" -Force
        $TicketsByResource = $timeEntries | Group-Object resourceID
        foreach ($Item in $TicketsByResource) {
            $Resource = $Resources | Where-Object id -in (($Item.name) )
            $item.group | Add-Member -NotePropertyName 'Resource' -NotePropertyValue $Resource.username -Force
            # return $Item.group
        }
  

        # identify any internal tickets
        $timeEntries | Add-Member -NotePropertyName 'InternalTicketBillableNormalHrs' -NotePropertyValue 0.0 -Force
        $timeEntries | Add-Member -NotePropertyName 'InternalTicketNonBillableNormalHrs' -NotePropertyValue 0.0 -Force
        $timeEntries | Add-Member -NotePropertyName 'InternalTicketBillableAftHrs' -NotePropertyValue 0.0 -Force
        $timeEntries | Add-Member -NotePropertyName 'InternalTicketNonBillableAftHrs' -NotePropertyValue 0.0 -Force
        $timeEntries | Add-Member -NotePropertyName 'InternalTicket' -NotePropertyValue 0.0 -Force


        $earliestDate = ($timeEntries | Measure-Object dateWorked -min).Minimum
        $CompanyTickets = Get-ATTickets -LastActionFromDate $earliestDate -CompanyIDs (29762985 , 0, 1, 29740186 , 29761818, 29762138, 29718567, 29762986)
   
        $InternalEntries = $timeEntries | Where-Object TicketID -in $CompanyTickets.id

        foreach ($i in $InternalEntries) {
            $items = $i | Where-Object { (($_.isNonBillable -eq $true) -or ($_.billingCodeID -in $ATnonBillableCodes)) }
            if ($items) {
                foreach ($item in $items) {
                    $item.InternalTicketNonBillableNormalHrs = $item.hoursWorked
                    $item.Internalticket = $item.hoursWorked
                } 
            }
            $items = $i | Where-Object { ($_.isNonBillable -ne $true) }
            if ($items) {
                foreach ($item in $items) {
                    $item.InternalTicketBillableNormalHrs = $item.hoursWorked
                    $item.Internalticket = $item.hoursWorked
                }   
            }
            #identify the afterhours billable
            $items = $i | Where-Object { ($_.billingCodeID -in $afterHrsBillingCodes) -and ($_.isNonBillable -ne $true) }
            if ($items) {
                foreach ($item in $items) {
                    $item.InternalTicketBillableAftHrs = $item.hoursWorked
                    $item.InternalTicketBillableNormalHrs = 0
                    $item.Internalticket = $item.hoursWorked
                }
            }
            #identify the afterhours nonbillable
            $items = $i | Where-Object { ($_.billingCodeID -in $afterHrsBillingCodes) -and ($_.isNonBillable -eq $true) }
            if ($items) {
                foreach ($item in $items) {
                    $item.InternalTicketNonBillableAftHrs = $item.hoursWorked
                    $item.InternalTicketNonBillableNormalHrs = 0
                    $item.Internalticket = $item.hoursWorked

                }
            }
        }


        #  NO Need to return a Value since the input object is alrerady modified (it is a reference object)
       # return $timeEntries
        return 
    }


    

    function Set-ATLogin() {
        <#
    .SYNOPSIS
    Allows automatic connection to the AutoTask API
    
    .DESCRIPTION
    Checks credentials and API integration code, then saves them encrypted within a file
    in the user's home\kiss-atapi path.
    Inline parameter values are accepted but not recommended — best practice is to leave
    them blank and enter values at the prompts, where they are handled as SecureStrings
    (not echoed to the screen and never stored in plain text).
    
    .PARAMETER l_username
    API username (usually an email address, NOT a firstname.lastname format).
    This is a globally usable Autotask API username.
    
    .PARAMETER l_pass
    Password for the API user.
    Saved as a DPAPI-encrypted string; never stored in plain text.
    
    .PARAMETER l_apiid
    API Integration Code for the API user.
    Saved as a DPAPI-encrypted string; never stored in plain text.
    
    .EXAMPLE
     Set-ATLogin
        there is already definition saved : for gokypolmtounjb6@KISSIT.CO.NZ
        If you wish to keep the old settings, then just hit return on that field without entering anything
        Enter a new API USER :
        now checking with the remote autotask API....
        will use the following autotask API intergface:   https://webservices6.autotask.net/ATServicesRest/
        Enter the USER's password (Alphanumerical and special):
        Enter the AT-API-ID  {alphanumerical}:
        Connection to the AutoaTask API was successfull: Your credentials work!, 
    
    .NOTES
    General notes
    #>
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $l_username, #= 'gokypolmtounjb6@KISSIT.CO.NZ'
            [string]
            $l_pass,
            [string]
            $l_apiid,
            [switch]$Force = $false
        )

        $saveobj = @{
            atapi    = ''
            UserName = ''
            Secret   = ''
            url      = ''
        }

        if (!(Test-Path -Path $kissATAPIpath)) {
            New-Item -Path $home -Name kiss-atapi -ItemType Directory | Out-Null
            Write-Host "Created a new directory: $kissATAPIpath"
        }
        else {
            if (Test-Path -Path "$kissATAPIpath\$kissATAPIfile") {
                $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
                if ($jsn) {
                    $r = $jsn | ConvertFrom-Json
                    # Only load existing values — do NOT print the raw JSON (it contains encrypted secrets)
                    Write-Host "Found existing saved credentials for: $($r.UserName)"
                }
                if ($r.url -and $r.secret -and $r.username -and $r.atapi) {
                    $saveobj = $r
                }
            }
        }

        # Allow inline parameters to pre-populate values (encrypted at rest immediately)
        if ($l_username) { $saveobj.UserName = $l_username }
        if ($l_pass) {
            $saveobj.Secret = ($l_pass | ConvertTo-SecureString -AsPlainText -Force) | ConvertFrom-SecureString
        }
        if ($l_apiid) {
            # Encrypt the API integration code the same way as the password
            $saveobj.atapi = ($l_apiid | ConvertTo-SecureString -AsPlainText -Force) | ConvertFrom-SecureString
        }

        write-verbose "Set-ATLogin: UserName = $($saveobj.UserName)"
        write-verbose "Set-ATLogin: Secret   = [encrypted — not logged]"
        write-verbose "Set-ATLogin: atapi    = [encrypted — not logged]"

        if ($saveobj.userName) {
            Write-Host "Existing saved login: $($saveobj.UserName)"
            Write-Host "Press Enter on any prompt to keep the existing value."
        }

        $i = Read-Host -Prompt "Enter API username (email address)"
        if ($i) { $saveobj.username = $i }

        Write-Host "Looking up Autotask zone for $($saveobj.username)..."
        $r = Invoke-RestMethod -Uri "http://webservices.autotask.net/atservicesrest/v1.0/zoneInformation?user=$($saveobj.username)"

        if ($r.url) {
            Write-Host "Autotask API endpoint: $($r.url)"
            $saveobj.url = $r.url
        }
        else {
            Write-Host "$($saveobj.username) is not recognised by the Autotask API, or the API could not be reached. Please check the username and try again." -ForegroundColor Red
            return
        }

        $i = Read-Host -Prompt "Enter the API user's password" -AsSecureString -ErrorAction SilentlyContinue
        if ($i.Length -gt 0) {
            $saveobj.Secret = $i | ConvertFrom-SecureString   # DPAPI-encrypt and store
        }

        # Prompt for API integration code as a SecureString so it is not echoed to the screen
        $i = Read-Host -Prompt "Enter the AT-API-ID" -AsSecureString -ErrorAction SilentlyContinue
        if ($i.Length -gt 0) {
            $saveobj.atapi = $i | ConvertFrom-SecureString    # DPAPI-encrypt and store
        }

        # Confirm credentials work before saving
        $testresult = Test-ATConnection -LoginInfo $saveobj
        if ($testresult) {
            Write-Host "Set-ATLogin: Connection verified — saving credentials." -ForegroundColor Green
        }
        elseif ($force -eq $true) {
            Write-Host "Set-ATLogin: Connection test failed but -Force was specified — saving anyway." -ForegroundColor Yellow
        }
        else {
            Write-Host "Set-ATLogin: Connection test failed — credentials not saved. Run again to retry." -ForegroundColor Red
            return
        }

        $jsn2 = ConvertTo-Json $saveobj
        # Note: $jsn2 contains encrypted strings only — safe to write to disk, but not logged
        Write-Verbose "Set-ATLogin: writing credentials file to $kissATAPIpath\$kissATAPIfile"
        Set-Content "$kissATAPIpath\$kissATAPIfile" -Value $jsn2
        Write-Host "Set-ATLogin: credentials saved successfully."

    }

    function Sync-AT365Calendar {
        <#
    .SYNOPSIS
    Synchronises Microsoft 365 calendar events with Autotask time entries using Microsoft Graph.

    .DESCRIPTION
    Reads time entries for the current engineer (ForMeOnly) going back to the Sunday of LastXWeeks weeks ago,
    then compares them against existing Microsoft 365 calendar events that contain an Autotask TimeEntry ID
    in their body.  For each time entry:
    - If a matching calendar event already exists, its start/end times are updated (and optionally its subject and body).
    - If no matching event is found, a new calendar event is created via New-AT365CalendarEvent.

    Requires the Microsoft Graph PowerShell SDK and Calendars.ReadWrite permissions.

    .PARAMETER LastXWeeks
    How many weeks back to retrieve time entries and check calendar events.
    Default is 1 (current week from the previous Sunday).

    .PARAMETER DoNotProvideDetaledInfoOnNew
    Switch. When set, newly created calendar events will have a minimal subject and body
    (no ticket title or company name lookup). Faster but less informative.

    .PARAMETER UpdateExistingInDetail
    Switch. When set, existing calendar events are also updated with full ticket and company detail
    (subject, company name, ticket title). Implies an extra Get-ATTickets call per entry.

    .EXAMPLE
    Sync-AT365Calendar -LastXWeeks 2

    .EXAMPLE
    Sync-AT365Calendar -LastXWeeks 2 -DoNotProvideDetaledInfoOnNew

    .EXAMPLE
    Sync-AT365Calendar -LastXWeeks 2 -UpdateExistingInDetail

    .NOTES
    Uses Get-AT365CalendarEvents internally to retrieve existing events.
    Calendar events are matched to time entries via the AutotaskTimeEntryID extended property.
    #>
        [CmdletBinding()]
        param (
            [Parameter(position=0)]

            [int]$LastXWeeks = 0,
            [switch]$DoNotProvideDetaledInfoOnNew = $false,
            [switch]$UpdateExistingInDetail = $false

        )
        $DateTostartCalCheck = Get-ATWeekStart -LastXWeeks $LastXWeeks 
        $timeEntries = Get-ATTimeEntries  -ForCalendar -ForMeOnly -FromDateLocal $DateTostartCalCheck  -IncludeSummaryNotes #-includeTicketDetails #-includeEngineerDetails -IncludeBillingDetails

        write-verbose "Sync-AT365Calendar: will check calendar items starting from $DateTostartCalCheck" #-ForegroundColor Green
        $foundCals = Get-AT365CalendarEvents -LastXWeeks ($LastXWeeks + 1) -OnlyGetItemWithAutotaskTimeEntryIDsInBody

        foreach ($timeEntry in $timeEntries) {
            $matchingCal = $foundCals | Where-Object { $_.AutotaskTimeEntryID -eq $timeEntry.id } | Select-Object -First 1

            if ((!$DoNotProvideDetaledInfoOnNew -eq $true) -or ($UpdateExistingInDetail -eq $true)) {
                $ticket = Get-ATTickets -id $timeEntry.ticketID -ForCalendar 
                $365Subject = "AT: " + $ticket.CompanyName + " - " + $ticket.Title
                $365Body = "TimeEntry:$($timeEntry.id)`nCompany: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)`nTicketNumber: $($ticket.ticketNumber)`nHours: $($timeEntry.hoursWorked)`nHoursToBill: $($timeEntry.hoursToBill)`nHoursWorked: $($timeEntry.hoursworked)`nisNOnBillable: $($timeEntry.isNonBillable)`nSyncd via Sync-AT365Calendar"   
            }
            else {
                $365Subject = "AT: Time Entry for TicketID $($timeEntry.ticketID)"
                $365Body = "TimeEntry:$($timeEntry.id)`nHours: $($timeEntry.hoursWorked)`nHoursToBill: $($timeEntry.hoursBillable)`nHoursWorked: $($timeEntry.hoursworked)`nisNOnBillable: $($timeEntry.isNonBillable)`nSynced via syncAT365Calendar"
            }
            if ($timeEntry.SummaryNotes) {
                $365Body = $365Body + "`n`nSummaryNotes: $($timeEntry.SummaryNotes)"
            }

            if ($MatchingCal) {
                write-verbose "Sync-AT365Calendar: Updating calendar item: Company: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)" #-ForegroundColor Cyan
                write-Host "`nUpdating calendar item: Company: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)" -ForegroundColor Cyan
                Update-AT365CalendarEvent -EventId $matchingCal.id -Subject $365Subject -Body $365Body -StartUTC $timeEntry.startDateTime -EndUTC $timeEntry.endDateTime        
            }
            else {
                write-verbose "Sync-AT365Calendar: Creating calendar item: Company: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)"
                write-Host "`nCreating calendar item: Company: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)" -ForegroundColor Blue
                New-AT365CalendarEvent -Subject $365Subject -Body $365Body -StartUTC $timeEntry.startDateTime -EndUTC $timeEntry.endDateTime -AutotaskTimeEntryID $timeEntry.id
            }

            Write-Host -NoNewline "."  
        }
        write-host "`nDone Sync-AT365Calendar" -ForegroundColor Green

    }




    function Sync-ATOutlookCalendar() {
        <#
    .SYNOPSIS
    Synchronises the local Outlook desktop calendar with Autotask time entries via COM automation.

    .DESCRIPTION
    Reads time entries for the current engineer (ForMeOnly) going back to the Sunday of LastXWeeks
    weeks ago, then compares them against existing Outlook calendar items that contain an Autotask
    TimeEntry ID in their body text.  For each time entry:
    - If a matching calendar item already exists, its start/end times are updated (and optionally its
      subject and body with full ticket detail).
    - If no matching item is found, a new Outlook appointment is created.

    Requires Outlook to be installed and configured on the local machine.
    Use Sync-AT365Calendar for Microsoft Graph / 365 online calendar instead.

    .PARAMETER LastXWeeks
    How many weeks back to retrieve time entries and check calendar items.
    Default is 1 (current week from the previous Sunday).

    .PARAMETER DoNotProvideDetaledInfoOnNew
    Switch. When set, newly created calendar items will have a minimal subject and body
    (no ticket title or company name lookup). Faster but less descriptive.

    .PARAMETER UpdateExistingInDetail
    Switch. When set, existing calendar items are also updated with full ticket and company detail
    (subject, company name, ticket title).

    .EXAMPLE
    Sync-ATOutlookCalendar -LastXWeeks 2

    .EXAMPLE
    Sync-ATOutlookCalendar -LastXWeeks 2 -DoNotProvideDetaledInfoOnNew

    .EXAMPLE
    Sync-ATOutlookCalendar -LastXWeeks 2 -UpdateExistingInDetail

    .NOTES
    Uses Outlook COM automation (New-Object -ComObject Outlook.Application).
    Calendar items are matched to time entries by searching the body for 'TimeEntry: <id>'.
    For Microsoft 365 / Graph-based calendar sync, use Sync-AT365Calendar instead.
    #>
        [CmdletBinding()]
        param (
            [Parameter(position=0)]
            [int]$LastXWeeks = 0,
            [switch]$DoNotProvideDetaledInfoOnNew = $false,
            [switch]$UpdateExistingInDetail = $false
            # [switch]$use365 = $false
        )
    
        $DateTostartCalCheck = Get-ATWeekStart -LastXWeeks $LastXWeeks 
        #$timeEntries = Get-ATTimeEntries  -ForCalendarLiteView -IncludeSummaryNotes -ForResourceID (Get-ATEngineerToReport) -FromDate $DateTostartCalCheck.ToUniversalTime() 
        $timeEntries = Get-ATTimeEntries  -ForCalendar -ForMeOnly -FromDateLocal $DateTostartCalCheck  -IncludeSummaryNotes #-includeTicketDetails -includeEngineerDetails -IncludeBillingDetails
        #$timeEntries = ($timeEntries |Where-Object {$null -ne $_.endDateTime} ) |Select-Object -Unique


        # $timeEntries

        # $timeEntries = Get-ATTimeEntries -LastXWeeks $LastXWeeks -ForCalendarLiteView -IncludeSummaryNotes        
        #$CURRENTDATE = GET-DATE -Hour 0 -Minute 0 -Second 0
       
        # $DateTostartCalCheck = $CURRENTDATE.AddDays(-7 * ($LastXWeeks + 1))
        write-verbose "Sync-ATOutlookCalendar: will check calendar items starting from $DateTostartCalCheck" #-ForegroundColor Green
        try {
            $Outlook = New-Object -ComObject Outlook.Application
            $Namespace = $Outlook.GetNamespace("MAPI")

            # Calendar folder (9 = olFolderCalendar)
            $Calendar = $Namespace.GetDefaultFolder(9)  
            $Items = $Calendar.Items   
        }
        catch {
            write-host "Sync-ATOutlookCalendar: sorry but there was an error connecting to Outlook. Please make sure you have Outlook installed and configured on this machine, and that you have access to the calendar." -ForegroundColor Red
            write-host "The error message was: $($_.Exception.Message)" -ForegroundColor Red
            return
        }

    

        $Items.Sort("[Start]")
        $Items.IncludeRecurrences = $false     # Optional: include recurring events
        $Items = $Items.Restrict("[Start] > '$($DateTostartCalCheck.ToString('g'))'")  # Only check events starting after this date
         
        write-verbose "Sync-ATOutlookCalendar: Checking $($items.count) calendar items for ExternalID user property..." #-ForegroundColor Green

        $foundCals = @()
        foreach ($cal in $items) {
            if ($cal.body -match 'TimeEntry:\s*(\d+)') {
                $number = $Matches[1]
                  
                Write-verbose "Sync-ATOutlookCalendar:Found calendar item.Body with TimeEntry: $([int]$number)  Subject: $($cal.Subject) Start: $($cal.Start) End: $($cal.End)" #-ForegroundColor Cyan
                $cal | Add-Member -NotePropertyName AutotaskTimeEntryID -NotePropertyValue ([int]$number) -Force
                $foundCals += $cal
            }
        }

        #$timeentriesCount = $timeEntries.Count


        foreach ($timeEntry in $timeEntries) {
            $matchingCal = $foundCals | Where-Object { $_.AutotaskTimeEntryID -eq $timeEntry.id } | Select-Object -first 1
            if ($matchingCal) {
                Write-verbose "Sync-ATOutlookCalendar: Updating calendar item: Company: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)" #-ForegroundColor Green
                $appt = $matchingCal
            }
            else {
                Write-host "`nSync-ATOutlookCalendar: Creating calendar item: Company: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)" -nonewline -ForegroundColor cyan
                $appt = $calendar.Items.Add()
            }
            # Used By Outlook code:  Add your external ID as a UserProperty
            #$prop = $appt.UserProperties.Add("AutotaskTimeEntryID", 1)   # 1 = Text field
            #$prop.Value = $timeEntry.id     # Your Autotask ID
                
            $appt.StartUTC = $timeEntry.startDateTime.toString('g')
            $appt.EndUTC = $timeEntry.endDateTime.toString('g')
            $appt.ReminderSet = $false
            if ((!$DoNotProvideDetaledInfoOnNew -eq $true) -or ($UpdateExistingInDetail -eq $true)) {
                $ticket = Get-ATTickets -id $timeEntry.ticketID -ForCalendar 
                $appt.Subject = "AT: " + $ticket.CompanyName + " - " + $ticket.Title
                $appt.Body = "TimeEntry:$($timeEntry.id)`nCompany: $($ticket.CompanyName)`nTicket Title: $($ticket.Title)`nTicketNumber: $($ticket.ticketNumber)`nHours: $($timeEntry.hoursWorked)`nHoursToBill: $($timeEntry.hoursBillable)`nHoursWorked: $($timeEntry.hoursworked)`nisNOnBillable: $($timeEntry.isNonBillable)`nCreated via Sync-ATOutlookCalendar"   
            }
            else {
                $appt.Subject = "AT: Time Entry for TicketID $($timeEntry.ticketID)"
                $appt.Body = "TimeEntry:$($timeEntry.id)`nHours: $($timeEntry.hoursWorked)`nHoursToBill: $($timeEntry.hoursBillable)`nHoursWorked: $($timeEntry.hoursworked)`nisNOnBillable: $($timeEntry.isNonBillable)`nCreated via Sync-ATOutlookCalendar"
            }
            if ($timeEntry.SummaryNotes) {
                $appt.Body = $appt.Body + "`n`nSummaryNotes: $($timeEntry.SummaryNotes)"
            }
            $appt.Save()
            Write-Host -NoNewline "."   
        }   
        write-Host "`nDone processing calendar items for $($timeEntries.count) time entries" -ForegroundColor Green
    }
  



    function Test-ATConnection {
        <#
    .SYNOPSIS
    Tests whether the saved (or supplied) Autotask API credentials are valid.

    .DESCRIPTION
    Attempts a live connection to the Autotask REST API Version endpoint using either
    credentials passed via the LoginInfo parameter or the credentials stored in the
    local kiss-atapi login file.

    Returns $true if the connection succeeds, or $null if it fails or credentials are missing.
    Used by Set-ATLogin to verify credentials before saving them.

    .PARAMETER LoginInfo
    Optional. A PSCustomObject containing the authentication details to test, with the
    properties: url, UserName, Secret (encrypted string), and atapi.
    If omitted, the saved login file is used.

    .PARAMETER LoginInfoPasswordAsPlainText
    Switch. When present, treats the Secret value in LoginInfo as plain text rather than
    a DPAPI-encrypted string. Only use this when passing credentials that have not yet
    been encrypted.

    .EXAMPLE
    # Test using the saved credentials
    Test-ATConnection

    .EXAMPLE
    # Test a specific login object
    $creds = [PSCustomObject]@{ url = "https://webservices6.autotask.net/ATServicesRest/"; UserName = "api@example.com"; Secret = $encryptedSecret; atapi = "MYAPIID" }
    Test-ATConnection -LoginInfo $creds

    .NOTES
    On success writes a green confirmation message to the host.
    On failure writes the exception message in yellow and returns $null.
    #>
        [CmdletBinding()]
        param(
            [PSCustomObject]$LoginInfo,
            [switch]$LoginInfoPasswordAsPlainText
        )
        # Build auth header via central helper (validates credentials exist and decrypts safely)
        $baseUrl = $null
        try {
            $kissATheader = Get-ATCredentialHeader -LoginInfo $LoginInfo -BaseUrl ([ref]$baseUrl)
        }
        catch {
            write-host "Test-ATConnection: Credentials are missing or could not be decrypted. Run Set-ATLogin first." -ForegroundColor Yellow
            return $null
        }
        write-verbose "Test-ATConnection: credentials loaded for URL $baseUrl, now testing connection"
        # Note: credential values are NOT logged — even at -Verbose — to avoid leaking secrets


        try {
            $versionUrl = "$baseUrl" + "v1.0/Version"
            Invoke-RestMethod -Method Get -Uri $versionUrl -Headers $kissATheader | Out-Null
            Write-host "Test-ATConnection: Connection to the Autotask API was successful — your credentials work!" -BackgroundColor Green
            return $true
        }
        catch {
            write-host "Test-ATConnection: Those credentials did not work." -ForegroundColor Yellow
            write-host "$($_.Exception.Message)" -ForegroundColor Yellow
            write-host "Please run Set-ATLogin again if you need to update your credentials." -ForegroundColor Yellow
            return $null
        }
    
    }

    function Test-ATWorkingDay() {
        <#
    .SYNOPSIS
    Determines whether a given date falls on a working day (true) or a weekend (false).

    .DESCRIPTION
    Returns $true when the supplied date is a weekday (Monday-Friday) and $false when
    it is a Saturday or Sunday. Public holidays are not currently accounted for.

    .PARAMETER date
    The DateTime value to evaluate.

    .EXAMPLE
    Test-ATWorkingDay -date (Get-Date)

    .EXAMPLE
    if (Test-ATWorkingDay -date $result.workDate) { "Working day" }

    .NOTES
    Public holiday exclusions are not implemented. Extend the function body with a
    holiday list if regional public holidays need to be treated as non-working days.
    #>
        [CmdletBinding()]
        param (
            [DateTime]$date
        )
        # Check if the day of the week is Saturday or Sunday
        if ($date.DayOfWeek -eq 'Saturday' -or $date.DayOfWeek -eq 'Sunday') {
            return $false
        }
        # Add any additional logic to exclude public holidays if needed
        # For example, you can maintain a list of public holidays and compare against it.
        # Otherwise, you can assume all weekdays are working days.
        return $true
    }


    function Update-AT365CalendarEvent {
        <#
    .SYNOPSIS
    Updates an existing Microsoft 365 calendar event via Microsoft Graph.

    .DESCRIPTION
    Connects to Microsoft Graph with Calendars.ReadWrite scope and updates the subject, body,
    start time, and end time of an existing calendar event identified by EventId.
    Reminders are disabled on save.

    StartUTC must be before EndUTC — the function exits with an error message if this is violated.

    .PARAMETER EventId
    The Microsoft Graph event ID of the calendar event to update.
    Obtain this from Get-AT365CalendarEvents or from a previous New-MgUserEvent call.

    .PARAMETER Subject
    The new subject line for the calendar event.

    .PARAMETER Body
    The new plain-text body content for the event.

    .PARAMETER StartUTC
    The new start date and time for the event, expressed in UTC.

    .PARAMETER EndUTC
    The new end date and time for the event, expressed in UTC. Must be after StartUTC.

    .EXAMPLE
    $events = Get-AT365CalendarEvents -LastXWeeks 1 -OnlyGetItemsWithAutotaskTimeEntryIDsInBody
    Update-AT365CalendarEvent -EventId $events[0].Id -Subject "Updated Subject" -Body "Updated body" -StartUTC (Get-Date).ToUniversalTime() -EndUTC (Get-Date).AddHours(2).ToUniversalTime()

    .NOTES
    Requires the Microsoft.Graph.Calendar module and Calendars.ReadWrite scope.
    Only the fields Subject, Body, Start, and End are updated — other event properties are preserved.
    #>
        [CmdletBinding()]
        param (
            [string]$EventId,
            [string]$Subject,
            [string]$Body,
            [DateTime]$StartUTC,
            [DateTime]$EndUTC   
        )
        if (-not(Get-Module -ListAvailable -Name  Microsoft.Graph)) { 
            # if (-not(Get-InstalledModule Microsoft.Graph)) { 
            #Get-Module -ListAvailable -Name Microsoft.Graph  
            Write-Host "Microsoft Graph module not found" -ForegroundColor Black -BackgroundColor Yellow
            $install = Read-Host "Do you want to install the Microsoft Graph Module?"
  
            if ($install -match "[yY]") {
                Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -AllowClobber -Force
            }
            else {
                Write-Host "Microsoft Graph module is required." -ForegroundColor Black -BackgroundColor Yellow
                throw "Microsoft Graph module is required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
            } 
        }
        Connect-MgGraph -Scopes  "Calendars.ReadWrite" | Out-Null

        $365me = (Get-MgContext).Account
        write-Host "Update-AT365CalendarEvent: connected to Microsoft Graph as $365me, now updating calendar event subject: $subject" -ForegroundColor Green
        if ($startUTC -ge $endUTC) {
            write-host "Update-AT365CalendarEvent: Error - StartUTC must be before EndUTC" -ForegroundColor Red
            return
        }

        $updatedEvent = Update-MgUserEvent -UserId $365me -EventId $EventId -BodyParameter @{
            subject      = $subject
            start        = @{
                dateTime = $startUTC.ToString("o") # ISO 8601 format for date-time
                timeZone = "UTC"
            }
            end          = @{
                dateTime = $EndUTC.ToString("o") # ISO 8601 format for date-time
                timeZone = "UTC"
            }
            isReminderOn = $false
            Body         = @{
                contentType = "Text"
                content     = $Body
            }
        } 
        if ($updatedEvent) {
            write-verbose "Update-AT365CalendarEvent: successfully updated event with ID '$EventId'"
        }
        else {
            write-host "Update-AT365CalendarEvent: failed to update event:  $subject" -ForegroundColor Red
        }
    }
  


function Get-ATCustomerReport {
    <#
    .SYNOPSIS
    Produces a time-entry report for one or more customers over a date range, summarised
    in chronological order (ascending or descending).

    .DESCRIPTION
    For each requested customer the function fetches all time entries within the specified
    date window and produces:

      1. Customer header  — company name, ID, and the reporting period.

      2. Time entries     — every entry posted within the window, sorted chronologically
                           and grouped by date. Each entry shows the engineer, ticket,
                           date, start/end times, hours, billing code, and summary notes.

      3. Per-ticket summary  — total and billable hours rolled up per ticket.

      4. Per-engineer summary — total and billable hours rolled up per engineer.

      5. Customer totals   — grand total hours, billable hours, non-billable hours, and
                             after-hours for the period.

    Customers can be identified by their Autotask CompanyID (numeric) or by a partial or
    exact CompanyName string.  Both parameters accept arrays so multiple customers can be
    reported in a single call.

    By default the date window is the whole of the previous calendar month
    (1st day 00:00 through last day 23:59 local time).  Supply -FromDate / -ToDate to
    override, or use -LastxMonths / -LastXWeeks for rolling windows.

    The function prints a formatted console report and returns a structured object whose
    properties (Header, TimeEntries, TicketSummary, EngineerSummary, Totals) can be
    piped to Export-Csv or used in further processing.

    .PARAMETER CompanyIDs
    One or more numeric Autotask company IDs.

    .PARAMETER CompanyNames
    One or more company name strings. Partial matches are accepted (the same
    -contains search used by Get-ATCompanies). Use -ExactNameMatch to require
    exact name matches.

    .PARAMETER ExactNameMatch
    Switch.  When set with -CompanyNames, only exact (case-insensitive) name matches
    are returned.

    .PARAMETER FromDate
    The local DateTime to use as the inclusive start of the reporting window.
    When omitted the first day of the previous calendar month (midnight) is used.

    .PARAMETER ToDate
    The local DateTime to use as the inclusive end of the reporting window.
    When omitted the last moment of the previous calendar month (23:59:59) is used.

    .PARAMETER LastxMonths
    Alternative: retrieve entries starting N calendar months ago from today.
    Overrides the default previous-month window but is overridden by -FromDate/-ToDate.

    .PARAMETER LastXWeeks
    Alternative: retrieve entries starting N weeks ago (from the previous Sunday).
    Overrides the default previous-month window but is overridden by -FromDate/-ToDate.

    .PARAMETER Descending
    Switch. Sorts time entries in descending date order (newest first).
    Default is ascending (oldest first).

    .PARAMETER SuppressConsoleOutput
    Switch. Suppresses the formatted console report; only the return object is produced.

    .PARAMETER OutputHtml
    Switch. Generates a self-contained HTML report.  Light mode (print-friendly) is used
    by default; add -DarkMode for the dark-themed version.

    .PARAMETER HtmlPath
    Full path for the HTML file.  Defaults to a timestamped file in $env:TEMP.

    .PARAMETER OpenInBrowser
    Switch. Generates the HTML and immediately opens it with Start-Process.
    Implies -OutputHtml.

    .PARAMETER DarkMode
    Switch. Uses the dark-themed CSS instead of the default light/print-friendly theme.

    .EXAMPLE
    # Previous month report for a customer by ID, printed to console
    Get-ATCustomerReport -CompanyIDs 29762985

    .EXAMPLE
    # Previous month for two customers by name, open as HTML in browser
    Get-ATCustomerReport -CompanyNames 'Acme','Globex' -OpenInBrowser

    .EXAMPLE
    # Custom date range, descending order, suppress console
    $r = Get-ATCustomerReport -CompanyIDs 29762985 `
             -FromDate (Get-Date '2026-01-01') `
             -ToDate   (Get-Date '2026-03-31') `
             -Descending -SuppressConsoleOutput
    $r | ForEach-Object {
        $_.TimeEntries | Export-Csv ".\$($_.Header.CompanyName)_TimeEntries.csv" -NoTypeInformation
    }

    .EXAMPLE
    # Last two months, dark-mode HTML saved to a specific path
    Get-ATCustomerReport -CompanyNames 'Smith %26 Sons' `
        -LastxMonths 2 -OutputHtml -DarkMode -HtmlPath 'C:\Reports\customer.html'

    .EXAMPLE
    # Exact name match, previous month, descending
    Get-ATCustomerReport -CompanyNames 'Acme Ltd' -ExactNameMatch -Descending

    .NOTES
    Time entries are fetched via Get-ATTimeEntries -FromDateLocal / -ToDate equivalents.
    Because Get-ATTimeEntries does not natively accept a ToDate filter the function
    post-filters the returned entries to the requested window.

    When -CompanyNames is used the function resolves names to IDs via Get-ATCompanies
    before fetching time entries.

    The HTML output reuses the same CSS themes as Get-ATTicketReport for visual
    consistency across reports.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByID')]
    param (
        [Parameter(ParameterSetName = 'ByID',   Mandatory = $true, Position = 0 , ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
        [Alias('id')]
        [int[]]$CompanyID,

        [Parameter(ParameterSetName = 'ByName', Mandatory = $true, Position = 0, ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
        [Alias('Name')]
        [Alias('Company')]
        [string[]]$CompanyName,

        [Parameter(ParameterSetName = 'ByName')]
        [switch]$ExactNameMatch = $false,

        # ── Date range ──────────────────────────────────────────────────────────
        [Nullable[DateTime]]$FromDate    = $null,
        [Nullable[DateTime]]$ToDate      = $null,
        [int]$LastxMonths                = 0,
        [int]$LastXWeeks                 = 0,

        # ── Output options ──────────────────────────────────────────────────────
        [switch]$Descending             = $false,
        [switch]$SuppressConsoleOutput  = $false,
        [switch]$OutputHtml             = $false,
        [string]$HtmlPath               = '',
        [switch]$OpenInBrowser          = $false,
        [switch]$DarkMode               = $false,
        [switch]$SuppressEntryDetails   = $false
    )

    # ── Inner helpers ────────────────────────────────────────────────────────────

    # Strip common HTML tags to plain text (identical pattern to Get-ATTicketReport)
    function ConvertFrom-ATHtml {
        param([string]$html)
        if (-not $html) { return '' }
        $text = $html `
            -replace '(?s)<br\s*/?>'   , "`n" `
            -replace '(?s)<p[^>]*>'    , "`n" `
            -replace '(?s)</p>'        , '' `
            -replace '(?s)<li[^>]*>'   , "`n  • " `
            -replace '(?s)<[^>]+>'     , '' `
            -replace '&amp;'           , '&' `
            -replace '&lt;'            , '<' `
            -replace '&gt;'            , '>' `
            -replace '&nbsp;'          , ' ' `
            -replace '&quot;'          , '"' `
            -replace '&#39;'           , "'" `
            -replace '(?m)^\s+$'       , '' `
            -replace "`n`n`n+"         , "`n`n"
        return $text.Trim()
    }

    # Write a section divider to console
    function Write-CRSection {
        param([string]$Title, [string]$Color = 'Cyan')
        if (-not $SuppressConsoleOutput) {
            Write-Host ''
            Write-Host "  ── $Title " -ForegroundColor $Color -NoNewline
            Write-Host ('─' * [Math]::Max(2, 60 - $Title.Length)) -ForegroundColor DarkGray
        }
    }

    # ── HTML builder ─────────────────────────────────────────────────────────────
    function Build-ATCustomerReportHtml {
        param(
            [System.Collections.Generic.List[PSCustomObject]]$Reports,
            [bool]$Dark = $false,
            [string]$PeriodLabel = ''
        )

        function hesc { param([string]$s); if (-not $s) { return '' }; $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' }

        $darkCss = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",Arial,sans-serif;font-size:14px;background:#1a1a2e;color:#e0e0e0;padding:24px}
h1{font-size:1.4rem;color:#c8d6e5;margin-bottom:4px}
.report-meta{font-size:.8rem;color:#666;margin-bottom:28px}
.customer{background:#16213e;border:1px solid #0f3460;border-radius:8px;margin-bottom:32px;overflow:hidden}
.c-head{background:#0f3460;padding:14px 18px}
.c-name{font-size:1.1rem;font-weight:600;color:#fff}
.c-sub{font-size:.82rem;color:#a0b4c8;margin-top:3px}
.hrs-bar{display:flex;flex-wrap:wrap;border-bottom:1px solid #0f3460;font-size:.82rem}
.hc{padding:7px 18px;border-right:1px solid #0f3460}
.hc:last-child{border-right:none}
.hl{font-size:.7rem;color:#7eb8f7;text-transform:uppercase}
.hv{font-weight:600;color:#e0e0e0}
.hv.ok{color:#4ade80}
.hv.nb{color:#f87171}
.hv.ah{color:#f7c948}
.sec{padding:12px 18px 0}
.sec-title{font-size:.72rem;text-transform:uppercase;letter-spacing:.08em;color:#7eb8f7;border-bottom:1px solid #0f3460;padding-bottom:5px;margin-bottom:10px}
.day-header{font-size:.8rem;font-weight:600;color:#7eb8f7;margin:10px 0 4px;padding:3px 0;border-bottom:1px solid #0f3460}
.te{background:#0d1b33;border-radius:4px;margin-bottom:6px;padding:7px 12px}
.te-hdr{display:flex;flex-wrap:wrap;gap:10px;align-items:baseline;margin-bottom:3px}
.te-time{color:#7eb8f7;font-size:.85rem}
.te-hrs{font-weight:600;color:#4ade80}
.te-hrs.nb{color:#f87171}
.te-hrs.ah{color:#f7c948}
.te-eng{color:#a0b4c8;font-size:.85rem}
.te-ticket{color:#7eb8f7;font-size:.8rem}
.te-bill{font-size:.76rem;color:#555;margin-bottom:2px}
.te-notes{color:#b0bec5;white-space:pre-wrap;font-size:.86rem;line-height:1.5}
.te-warn{font-size:.76rem;color:#f7c948;margin-top:2px}
table.summ{width:100%;border-collapse:collapse;font-size:.84rem;margin-bottom:12px}
table.summ th{text-align:left;color:#7eb8f7;font-weight:500;padding:4px 10px 5px 0;border-bottom:1px solid #0f3460}
table.summ td{padding:5px 10px 5px 0;border-bottom:1px solid #0a1628;color:#c8d6e5}
table.summ td.r{text-align:right;padding-right:18px}
table.summ td.ok{color:#4ade80}
table.summ td.nb{color:#f87171}
.summ-box{background:#16213e;border:1px solid #0f3460;border-radius:8px;padding:14px 18px;margin-bottom:24px}
.summ-box h2{font-size:.8rem;text-transform:uppercase;letter-spacing:.08em;color:#7eb8f7;margin-bottom:10px}
'@

        $lightCss = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",Arial,sans-serif;font-size:13px;background:#fff;color:#111;padding:20px}
h1{font-size:1.3rem;color:#1a1a2e;margin-bottom:3px}
.report-meta{font-size:.78rem;color:#777;margin-bottom:22px}
@media print{body{padding:8px}.customer{page-break-inside:avoid;margin-bottom:18px}}
.customer{text-align:Left;border:1px solid #bbb;border-radius:6px;margin-bottom:24px;overflow:hidden}
.c-head{background:#e8edf2;padding:10px 16px;border-bottom:1px solid #bbb}
.c-name{font-size:1.05rem;font-weight:700;color:#111}
.c-sub{font-size:.8rem;color:#444;margin-top:2px}
.hrs-bar{display:flex;flex-wrap:wrap;border-bottom:1px solid #ccc;font-size:.8rem}
.hc{padding:5px 14px;border-right:1px solid #ddd}
.hc:last-child{border-right:none}
.hl{font-size:.67rem;color:#3a5a8c;text-transform:uppercase}
.hv{font-weight:700;color:#111}
.hv.ok{color:#1a7a3a}
.hv.nb{color:#c0000a}
.hv.ah{color:#b06000}
.sec{padding:10px 16px 0}
.sec-title{font-size:.69rem;text-transform:uppercase;letter-spacing:.07em;color:#3a5a8c;border-bottom:1px solid #ccc;padding-bottom:4px;margin-bottom:8px}
.day-header{font-size:.78rem;font-weight:700;color:#3a5a8c;margin:8px 0 3px;padding:2px 0;border-bottom:1px solid #ddd}
.te{border:1px solid #eee;border-radius:3px;margin-bottom:5px;padding:6px 10px;background:#fdfdfd}
.te-hdr{display:flex;flex-wrap:wrap;gap:8px;align-items:baseline;margin-bottom:2px}
.te-time{color:#3a5a8c;font-size:.83rem}
.te-hrs{font-weight:700;color:#1a7a3a}
.te-hrs.nb{color:#c0000a}
.te-hrs.ah{color:#b06000}
.te-eng{color:#555;font-size:.83rem}
.te-ticket{color:#3a5a8c;font-size:.78rem}
.te-bill{font-size:.73rem;color:#888;margin-bottom:2px}
.te-notes{color:#333;white-space:pre-wrap;font-size:.85rem;line-height:1.5}
.te-warn{font-size:.73rem;color:#b06000;margin-top:2px}
table.summ{width:100%;border-collapse:collapse;font-size:.83rem;margin-bottom:10px}
table.summ th{text-align:left;color:#3a5a8c;font-weight:600;padding:3px 10px 4px 0;border-bottom:1px solid #ccc}
table.summ td{padding:4px 10px 4px 0;border-bottom:1px solid #eee;color:#111}
table.summ td.r{text-align:right;padding-right:18px}
table.summ td.ok{color:#1a7a3a}
table.summ td.nb{color:#c0000a}
.summ-box{border:1px solid #bbb;border-radius:6px;padding:12px 16px;margin-bottom:20px}
.summ-box h2{font-size:.76rem;text-transform:uppercase;letter-spacing:.07em;color:#3a5a8c;margin-bottom:8px}
'@

        $css   = if ($Dark) { $darkCss } else { $lightCss }
        $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        $blocks = [System.Text.StringBuilder]::new()

        # ── Grand summary table (all customers) ───────────────────────────────
        if ($Reports.Count -gt 1) {
            $null = $blocks.Append("<div class='summ-box'><h2>Report Summary — $($Reports.Count) customers — $([System.Net.WebUtility]::HtmlEncode($PeriodLabel))</h2>")
            $null = $blocks.Append("<table class='summ'><tr><th>Customer</th><th class='r'>Total hrs</th><th class='r'>Billable</th><th class='r'>Non-Bill</th><th class='r'>Entries</th></tr>")
            foreach ($rpt in $Reports) {
                $nbClass = if ($rpt.Totals.NonBillableHours -gt 0) { ' nb' } else { ' ok' }
                $null = $blocks.Append("<tr><td>$(hesc $rpt.Header.CompanyName)</td><td class='r'>$($rpt.Totals.TotalHours)</td><td class='r ok'>$($rpt.Totals.BillableHours)</td><td class='r$nbClass'>$($rpt.Totals.NonBillableHours)</td><td class='r'>$($rpt.Totals.TimeEntryCount)</td></tr>")
            }
            $null = $blocks.Append('</table></div>')
        }

        # ── Per-customer blocks ───────────────────────────────────────────────
        foreach ($rpt in $Reports) {
            $hdr = $rpt.Header
            $null = $blocks.Append("<div class='customer'>")
            $null = $blocks.Append("<div class='c-head'><div class='c-name'>$(hesc $hdr.CompanyName)</div><div class='c-sub'>Period: $(hesc $hdr.Period) &nbsp;|&nbsp; Company ID: $($hdr.CompanyID)</div></div>")

            # Hours bar
            $nbClass = if ($rpt.Totals.NonBillableHours -gt 0) { ' nb' } else { '' }
            $ahClass = if ($rpt.Totals.AfterHours -gt 0)       { ' ah' } else { '' }
            $null = $blocks.Append("<div class='hrs-bar'>")
            $null = $blocks.Append("<div class='hc'><div class='hl'>Total hrs</div><div class='hv'>$($rpt.Totals.TotalHours)</div></div>")
            $null = $blocks.Append("<div class='hc'><div class='hl'>Billable</div><div class='hv ok'>$($rpt.Totals.BillableHours)</div></div>")
            $null = $blocks.Append("<div class='hc'><div class='hl'>Non-Billable</div><div class='hv$nbClass'>$($rpt.Totals.NonBillableHours)</div></div>")
            if ($rpt.Totals.AfterHours -gt 0) {
                $null = $blocks.Append("<div class='hc'><div class='hl'>After Hours</div><div class='hv$ahClass'>$($rpt.Totals.AfterHours)</div></div>")
            }
            $null = $blocks.Append("<div class='hc'><div class='hl'>Entries</div><div class='hv'>$($rpt.Totals.TimeEntryCount)</div></div>")
            $null = $blocks.Append("</div>")

            # Engineer summary table
            if ($rpt.EngineerSummary.Count -gt 0) {
                $null = $blocks.Append("<div class='sec'><div class='sec-title'>Engineer Summary</div>")
                $null = $blocks.Append("<table class='summ'><tr><th>Engineer</th><th class='r'>Total hrs</th><th class='r'>Billable</th><th class='r'>After Hrs</th></tr>")
                foreach ($eng in $rpt.EngineerSummary) {
                    $null = $blocks.Append("<tr><td>$(hesc $eng.Engineer)</td><td class='l'>$($eng.TotalHours)</td><td class='l ok'>$($eng.BillableHours)</td><td class='l ah'>$(if($eng.AfterHours -gt 0){$eng.AfterHours}else{'—'})</td></tr>")
                }
                $null = $blocks.Append('</table></div>')
            }

            # Ticket summary table
            if ($rpt.TicketSummary.Count -gt 0) {
                $null = $blocks.Append("<div class='sec'><div class='sec-title'>Ticket Summary</div>")
                $null = $blocks.Append("<table class='summ'><tr><th>Ticket</th><th class='l'>Title</th><th class='l'>Total hrs</th><th class='l'>Billable</th></tr>")
                foreach ($tk in $rpt.TicketSummary) {
                    $null = $blocks.Append("<tr><td>$(hesc $tk.TicketNumber)</td><td>$(hesc ($tk.TicketTitle -replace '(.{50}).+','$1…'))</td><td class='l'>$($tk.TotalHours)</td><td class='l ok'>$($tk.BillableHours)</td></tr>")
                }
                $null = $blocks.Append('</table></div>')
            }

            # if ($SuppressEntryDetails) {
            #     $null = $blocks.Append("<div class='sec'><div class='sec-title'>Time Entries</div><div>(details suppressed)</div></div>")
            #     #$null = $blocks.Append('</div>')  # close .customer
            #    # continue
            # }


            # Time entries grouped by day
            $null = $blocks.Append("<div class='sec'><div class='sec-title'>Time Entries ($($rpt.TimeEntries.Count))</div>")
            $lastDay = $null
            If ($SuppressEntryDetails) {    
                $null = $blocks.Append("<div>(details suppressed)</div>")  
            }
            else {
             foreach ($te in $rpt.TimeEntries) {
                $day = ($te.DateWorked -replace 'T.*', '')
                if ($day -ne $lastDay) {
                    if ($lastDay) { $null = $blocks.Append('</div>') }
                    $null = $blocks.Append("<div class='day-header'>$(hesc $day)</div><div>")
                    $lastDay = $day
                }
                $hrsClass = if (-not $te.IsBillable) { ' nb' } elseif ($te.AfterHours -gt 0) { ' ah' } else { '' }
                $timeStr  = if ($te.HasTimestamp) { "$(hesc $te.StartTime)–$(hesc $te.EndTime)" } else { '(no time)' }
                $billTag  = if (-not $te.IsBillable) { ' [NB]' } else { '' }
                $null = $blocks.Append("<div class='te'>")
                $null = $blocks.Append("<div class='te-hdr'><span class='te-time'>$timeStr</span><span class='te-hrs$hrsClass'>$($te.HoursWorked)h$billTag</span><span class='te-eng'>$(hesc $te.Engineer)</span>")
                if ($te.TicketNumber) { $null = $blocks.Append("<span class='te-ticket'>$(hesc $te.TicketNumber) | $(hesc $te.TicketTitle)</span>") }
                $null = $blocks.Append('</div>')
                $null = $blocks.Append("<div class='te-bill'>$(hesc $te.BillingCode)</div>")
                if (-not $te.HasTimestamp) { $null = $blocks.Append("<div class='te-warn'>⚠ No clock start/end recorded</div>") }
                if ($te.SummaryNotes) { $null = $blocks.Append("<div class='te-notes'>$(hesc $($te.SummaryNotes -replace "(\r?\n){2,}", "`r`n"))</div>") }
                $null = $blocks.Append('</div>')
            }
        }


            if ($lastDay) { $null = $blocks.Append('</div>') }  # close last day wrapper
            $null = $blocks.Append('</div>')  # close sec

            $null = $blocks.Append('</div>')  # close .customer
        }

        return @"
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Customer Time Report</title>
<style>$css</style>
</head>
<body>
<h1>Customer Time Report</h1>
<div class='report-meta'>Generated $stamp &nbsp;|&nbsp; Period: $(hesc $PeriodLabel)</div>
$($blocks.ToString())
</body>
</html>
"@
    }

    # ════════════════════════════════════════════════════════════════════════════
    # ── 1. Resolve the reporting date window ────────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════════

    $today = Get-Date -Hour 0 -Minute 0 -Second 0

    if ($null -ne $FromDate -and $null -ne $ToDate) {
        # Explicit window supplied
        $windowFrom = $FromDate
        $windowTo   = $ToDate.Date.AddDays(1).AddSeconds(-1)   # end-of-day
    }
    elseif ($LastxMonths -gt 0) {
        $windowFrom = $today.AddMonths(-$LastxMonths)
        $windowTo   = $today.AddSeconds(-1)
    }
    elseif ($LastXWeeks -gt 0) {
        $windowFrom = Get-ATWeekStart -LastXWeeks $LastXWeeks
        $windowTo   = $today.AddSeconds(-1)
    }
    else {
        # Default: the whole of the previous calendar month
        $firstOfThisMonth = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
        $windowFrom = $firstOfThisMonth.AddMonths(-1)
        $windowTo   = $firstOfThisMonth.AddSeconds(-1)
    }

    $periodLabel = "$($windowFrom.ToString('yyyy-MM-dd'))  to  $($windowTo.ToString('yyyy-MM-dd'))"
    Write-Host "Get-ATCustomerReport: reporting period $periodLabel" -ForegroundColor Cyan

    # ════════════════════════════════════════════════════════════════════════════
    # ── 2. Resolve company IDs ───────────────────────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════════

    $resolvedCompanies = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($PSCmdlet.ParameterSetName -eq 'ByName') {
        write-host "Get-ATCustomerReport: resolving company names $($CompanyName -join ", ") to IDs..." -ForegroundColor Cyan
        foreach ($name in $CompanyName) {
            Write-Host "Get-ATCustomerReport: resolving company name '$name'..." -ForegroundColor DarkCyan
            $matches_ = Get-ATCompanies -CompanyName $name -exactNameMatch:$ExactNameMatch
            if (-not $matches_) {
                Write-Warning "Get-ATCustomerReport: no company found matching '$name' — skipping."
                continue
            }
            foreach ($c in $matches_) { $resolvedCompanies.Add($c) }
        }
    }
    else {
        foreach ($cid in $CompanyID) {
            Write-Host "Get-ATCustomerReport: looking up company ID $cid..." -ForegroundColor DarkCyan
            $c = Get-ATCompanies -id $cid | Select-Object -First 1
            if (-not $c) {
                Write-Warning "Get-ATCustomerReport: company ID $cid not found — skipping."
                continue
            }
            $resolvedCompanies.Add($c)
        }
    }

    if ($resolvedCompanies.Count -eq 0) {
        write-host "Get-ATCustomerReport: no valid companies found  via approach $($PSCmdlet.ParameterSetName)." -ForegroundColor Red
        Write-Warning 'Get-ATCustomerReport: no companies resolved — nothing to report.'
        return
    }

    # ════════════════════════════════════════════════════════════════════════════
    # ── 3. Pre-fetch shared lookup tables ───────────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════════

    $billingCodes   = Get-ATBillingCodes
    #$allEngineers   = Get-ATEngineers


            Write-Host "Get-ATCustomerReport: fetching ALL time entries for during that time windows - since the autotask timeentry identfies the ticket not the company" -ForegroundColor Cyan

        # ── 4a. Pull time entries from the API (from windowFrom onward) ─────────
        #        then post-filter to windowTo because Get-ATTimeEntries has no ToDate
        $AllrawEntries = Get-ATTimeEntries `
            -FromDateLocal     $windowFrom `
            -IncludeSummaryNotes  -UntilDateLocal $($windowTo.ToString('yyyy-MM-ddT23:59:59')) 
    #write-host $AllrawEntries.Count " total entries fetched from API for the entire date range (before post-filtering to specific customers)." -ForegroundColor DarkGray

#     $AllrawEntries
#   return
    # ════════════════════════════════════════════════════════════════════════════
    # ── 4. Build a report object per customer ───────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════════

    $reports = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($company in $resolvedCompanies) {

        $acompanyName = $company.companyName
        $acompanyID   = $company.id
        Write-Verbose "Processing company '$acompanyName' (ID $acompanyID)..."
        # # Post-filter: keep only entries belonging to this customer within the window
        # # The CompanyID on a time entry is populated by Get-ATTimeEntries via ticket lookup
        $rawEntries = $AllrawEntries | Where-Object {
            $_.CompanyID -eq $acompanyID } # -and $_.dateWorked -le $windowTo.ToString('yyyy-MM-ddT23:59:59')        }

        $entryCount = if ($rawEntries) { @($rawEntries).Count } else { 0 }
        Write-Host "  → $entryCount time entr$(if($entryCount -ne 1){'ies'}else{'y'}) found within timeframe." -ForegroundColor DarkGray

        # ── 4b. Sort entries chronologically (or reverse) ───────────────────────
        $sortedEntries = if ($Descending) {
            @($rawEntries | Sort-Object dateWorked -Descending)
        }
        else {
            @($rawEntries | Sort-Object dateWorked)
        }

        # ── 4c. Build the normalised entry row list ─────────────────────────────
        $entryRows = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($e in $sortedEntries) {
            $hasTimestamp  = ($null -ne $e.startDateTime) -and ($null -ne $e.endDateTime)
            $engineerLabel = if ($e.Engineer)      { $e.Engineer }      else { "ID:$($e.resourceID)" }
            $bcName        = if ($e.BillingCode)   { $e.BillingCode }   `
                             else { ($billingCodes | Where-Object id -eq $e.billingCodeID | Select-Object -First 1).name }

            $startDisplay = if ($hasTimestamp -and $e.startDateTimeLocal) {
                $e.startDateTimeLocal.ToString('HH:mm')
            } else { '' }
            $endDisplay   = if ($hasTimestamp -and $e.endDateTimeLocal) {
                $e.endDateTimeLocal.ToString('HH:mm')
            } else { '' }

            $entryRows.Add([PSCustomObject]@{
                CompanyID      = $acompanyID
                CompanyName    = $acompanyName
                TimeEntryID    = $e.id
                TicketID       = $e.ticketID
                TicketNumber   = if ($e.TicketNumber) { $e.TicketNumber } else { '' }
                TicketTitle    = if ($e.Title)  { $e.Title  } else { '' }
                DateWorked     = ($e.dateWorked -replace 'T.*', '')
                StartTime      = $startDisplay
                EndTime        = $endDisplay
                HasTimestamp   = $hasTimestamp
                Engineer       = $engineerLabel
                HoursWorked    = [Math]::Round($e.hoursWorked, 2)
                AfterHours     = $e.afterHours
                IsBillable     = -not ($e.isNonBillable -eq $true)
                BillingCode    = $bcName
                SummaryNotes   = if ($e.summaryNotes) { $e.summaryNotes } else { '' }
                NonBillableHrs = $e.hoursNonBillable
            })
        }


        # ── 4d. Roll-up totals ───────────────────────────────────────────────────
        $totalHrs       = [Math]::Round(($entryRows | Measure-Object HoursWorked    -Sum).Sum, 2)
        $billableHrs    = [Math]::Round(($entryRows | Where-Object IsBillable       | Measure-Object HoursWorked -Sum).Sum, 2)
        $nonBillableHrs = [Math]::Round(($entryRows | Measure-Object NonBillableHrs -Sum).Sum, 2)
        $afterHrs       = [Math]::Round(($entryRows | Measure-Object AfterHours     -Sum).Sum, 2)
        $noStampHrs     = [Math]::Round(($entryRows | Where-Object { -not $_.HasTimestamp } | Measure-Object HoursWorked -Sum).Sum, 2)

        $totals = [PSCustomObject]@{
            TotalHours      = $totalHrs
            BillableHours   = $billableHrs
            NonBillableHours= $nonBillableHrs
            AfterHours      = $afterHrs
            NoTimestampHours= $noStampHrs
            TimeEntryCount  = $entryRows.Count
        }

        # ── 4e. Per-ticket summary ───────────────────────────────────────────────
        $ticketSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
        $entryRows | Group-Object TicketID | ForEach-Object {
            $grp        = $_.Group
            $firstEntry = $grp | Select-Object -First 1
            $ticketSummary.Add([PSCustomObject]@{
                TicketID      = $_.Name
                TicketNumber  = $firstEntry.TicketNumber
                TicketTitle   = $firstEntry.TicketTitle
                TotalHours    = [Math]::Round(($grp | Measure-Object HoursWorked -Sum).Sum, 2)
                BillableHours = [Math]::Round(($grp | Where-Object IsBillable | Measure-Object HoursWorked -Sum).Sum, 2)
                EntryCount    = $grp.Count
            })

        }

        # ── 4f. Per-engineer summary ─────────────────────────────────────────────
        $engineerSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
        $entryRows | Group-Object Engineer | ForEach-Object {
            $grp = $_.Group
            $engineerSummary.Add([PSCustomObject]@{
                Engineer      = $_.Name
                TotalHours    = [Math]::Round(($grp | Measure-Object HoursWorked -Sum).Sum, 2)
                BillableHours = [Math]::Round(($grp | Where-Object IsBillable | Measure-Object HoursWorked -Sum).Sum, 2)
                AfterHours    = [Math]::Round(($grp | Measure-Object AfterHours -Sum).Sum, 2)
                EntryCount    = $grp.Count
            })
        }

        # ── 4g. Console output ───────────────────────────────────────────────────
        if (-not $SuppressConsoleOutput) {
            $sortLabel = if ($Descending) { 'newest first' } else { 'oldest first' }

            Write-Host ''
            Write-Host ('═' * 70) -ForegroundColor DarkGray
            Write-Host "  CUSTOMER   $acompanyName" -ForegroundColor Yellow
            Write-Host "  Period     $periodLabel   ($sortLabel)" -ForegroundColor Yellow
            Write-Host ('═' * 70) -ForegroundColor DarkGray

            $billColor = if ($billableHrs -lt $totalHrs) { 'DarkYellow' } else { 'Green' }
            Write-Host "  Total Hours     : $totalHrs" -ForegroundColor White
            Write-Host "  Billable Hours  : $billableHrs"  -ForegroundColor $billColor
            if ($nonBillableHrs -gt 0) {
                Write-Host "  Non-Billable    : $nonBillableHrs" -ForegroundColor Red
            }
            if ($afterHrs -gt 0) {
                Write-Host "  After Hours     : $afterHrs"  -ForegroundColor DarkYellow
            }
            if ($noStampHrs -gt 0) {
                Write-Host "  (incl. $noStampHrs hrs with no clock-in/out time)" -ForegroundColor DarkYellow
            }

            # Engineer summary
            Write-CRSection 'Engineer Summary'
            foreach ($eng in $engineerSummary | Sort-Object TotalHours -Descending) {
                $ahSuffix = if ($eng.AfterHours -gt 0) { "  (incl. $($eng.AfterHours)h after-hours)" } else { '' }
                Write-Host "    $($eng.Engineer.PadRight(30))  $($eng.TotalHours)h  (Billable: $($eng.BillableHours)h)$ahSuffix" -ForegroundColor Gray
            }

            # Ticket summary
            if ($ticketSummary.Count -gt 0) {
                Write-CRSection "Ticket Summary ($($ticketSummary.Count))"
                $ticketSummary | Sort-Object TotalHours -Descending | ForEach-Object {
                    $tkLabel = if ($_.TicketNumber) { $_.TicketNumber } else { "(task/no ticket)" }
                    $tkTitle = if ($_.TicketTitle) { $_.TicketTitle -replace '(.{35}).+','$1…' } else { '' }
                    Write-Host ("    {0,-20}  {1,-37}  {2,5}h  bill:{3}h" -f $tkLabel, $tkTitle, $_.TotalHours, $_.BillableHours) -ForegroundColor Gray
                }
            }

            if ($SuppressEntryDetails) {
                Write-Host ''
                Write-Host '  (time entry details suppressed)' -ForegroundColor DarkGray
                Write-Host ('═' * 70) -ForegroundColor DarkGray
                continue
            }
            # Detailed time entries
            Write-CRSection "Time Entries ($($entryRows.Count))"
            if ($entryRows.Count -gt 0) {
                $lastDay = $null
                foreach ($te in $entryRows) {
                    # Day separator
                    if ($te.DateWorked -ne $lastDay) {
                        Write-Host ''
                        Write-Host "    ▸ $($te.DateWorked)" -ForegroundColor Cyan
                        $lastDay = $te.DateWorked
                    }

                    $eColor  = if ($te.AfterHours -gt 0)     { 'DarkYellow' } `
                               elseif ($te.NonBillableHrs -gt 0) { 'Red'     } `
                               else                             { 'DarkGray'  }
                    $timeStr = if ($te.HasTimestamp) { "$($te.StartTime)–$($te.EndTime)" } else { '(no time⚠)' }
                    $billTag = if ($te.IsBillable)   { '' } else { ' [NB]' }
                    $tkStr   = if ($te.TicketNumber) { "  [$($te.TicketNumber)]" } else { '' }

                    Write-Host "      $timeStr  $($te.HoursWorked)h  $($te.Engineer)$billTag$tkStr $($te.TicketTitle)" -ForegroundColor White
                    Write-Host "      Billing: $($te.BillingCode)" -ForegroundColor $eColor
                    if ($te.SummaryNotes) {
                        ($te.SummaryNotes -replace "(\r?\n){2,}", "`n") -split "`n" | ForEach-Object {
                            Write-Host "        $_" -ForegroundColor Gray
                        }
                    }
                }
                if ($noStampHrs -gt 0) {
                    Write-Host ''
                    Write-Host "    ⚠  $noStampHrs hrs recorded without a clock start/end time" -ForegroundColor DarkYellow
                }
            }
            else {
                Write-Host '    (no time entries in this period)' -ForegroundColor DarkGray
            }

            Write-Host ''
            Write-Host ('═' * 70) -ForegroundColor DarkGray
        }

        # ── 4h. Accumulate into report list ─────────────────────────────────────
        $reports.Add([PSCustomObject]@{
            Header          = [PSCustomObject]@{
                CompanyID   = $acompanyID
                CompanyName = $acompanyName
                Period      = $periodLabel
                FromDate    = $windowFrom
                ToDate      = $windowTo
                SortOrder   = if ($Descending) { 'Descending' } else { 'Ascending' }
            }
            TimeEntries     = $entryRows.ToArray()
            TicketSummary   = $ticketSummary.ToArray()
            EngineerSummary = $engineerSummary.ToArray()
            Totals          = $totals
        })
    }

    # ════════════════════════════════════════════════════════════════════════════
    # ── 5. Grand rollup when multiple customers ──────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════════

    if ($reports.Count -gt 1 -and -not $SuppressConsoleOutput) {
        $grandTotal    = [Math]::Round(($reports | ForEach-Object { $_.Totals.TotalHours    } | Measure-Object -Sum).Sum, 2)
        $grandBillable = [Math]::Round(($reports | ForEach-Object { $_.Totals.BillableHours } | Measure-Object -Sum).Sum, 2)

        Write-Host ''
        Write-Host ('─' * 70) -ForegroundColor DarkGray
        Write-Host "  REPORT SUMMARY — $($reports.Count) customers" -ForegroundColor White
        Write-Host "  Period         : $periodLabel"       -ForegroundColor White
        Write-Host "  Grand Total hrs: $grandTotal"        -ForegroundColor White
        Write-Host "  Grand Billable : $grandBillable"     -ForegroundColor Green
        Write-Host ('─' * 70) -ForegroundColor DarkGray

        $reports | Sort-Object { $_.Header.CompanyName } | ForEach-Object {
            Write-Host ("  {0,-35}  {1,6}h  bill:{2}h  entries:{3}" -f `
                $_.Header.CompanyName, $_.Totals.TotalHours, $_.Totals.BillableHours, $_.Totals.TimeEntryCount) `
                -ForegroundColor Gray
        }
        Write-Host ''
    }

    # ════════════════════════════════════════════════════════════════════════════
    # ── 6. HTML output ───────────────────────────────────────────────────────────
    # ════════════════════════════════════════════════════════════════════════════

    if ($OutputHtml -or $OpenInBrowser) {
        $htmlContent = Build-ATCustomerReportHtml -Reports $reports -Dark:$DarkMode -PeriodLabel $periodLabel
        if (-not $HtmlPath) {
            $stamp    = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $HtmlPath = Join-Path $env:TEMP "ATCustomerReport_$stamp.html"
        }
        $htmlContent | Out-File -FilePath $HtmlPath -Encoding UTF8 -Force
        Write-Host "Get-ATCustomerReport: HTML saved to $HtmlPath" -ForegroundColor Cyan
        if ($OpenInBrowser) { Start-Process $HtmlPath }
    }

    # Return the array of report objects (one per customer)
    return $reports.ToArray()
}



# SIG # Begin signature block
# MIIFgwYJKoZIhvcNAQcCoIIFdDCCBXACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCMss8FCF3A3Gay
# oqMKa89ZEUin6bhp9GB4m1zm2iJ62aCCAv4wggL6MIIB4qADAgECAhA7Wkn363I8
# sU+TMwSw335wMA0GCSqGSIb3DQEBCwUAMBUxEzARBgNVBAMMClNlYW4gTWFjZXkw
# HhcNMjYwNDA0MjIxNTM5WhcNMzYwNDA0MjIyNTM5WjAVMRMwEQYDVQQDDApTZWFu
# IE1hY2V5MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6/RB8ks87nu6
# LqUgWXq02tdjYt427XKkEW7vFYFqr35woedz7nUwIgMcyDmbiTtOdzDAFJl4ld3/
# TJEVeyndCqePz+LsXRBk3nDxhouuh+ORnyn7ga3FFwp7jSmTiTr/LWMy8gZqhsvU
# sBCQWPA6OaJy8x0iGAjkKqWjwiO8lepPHR9MeTuRsiVI0GYbxdyf+2If8Lhhqq7R
# BwaNhTTvYjDGG95VaaIOngPYxDnz1UsWjLiCA0vrq+ZEeiT/gOvtAzRrH6NMZHVE
# JekVhuByAreI9StjTwyzmiIwZhK95vwHVaXpF4OXFzSpneGihJPeoU/M9PToeJnm
# EHw7rIWQrQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwHQYDVR0OBBYEFC3Zqjwt22ikPMgP/7MK7ULso51/MA0GCSqGSIb3DQEB
# CwUAA4IBAQB6jHCzFEeK/R1TwNZotJtRIJX67GTeQvY/LmLuLCo74td0rzIUddp/
# DmZWszqlNnEcnNkRnWJ1A07ge7FTn51biripsHSxX7f6xfSc/5HbcUm9diANjYXV
# 18hEeSc2E0Yw2Xz1HI35owaQZotWZX9I7CKLiCXfOEEtWgbS/+Ff7PxQ7C60zwP+
# OSmthwdUeSeDPnSr6IXnTQ0/DKlEMW1wFfhinGvT20J/dJQxlm66vE4WfKDrrDln
# TAQaVWe5CvhZ1q84AV8o5zz13mO3HWJ+2+2bqj1+CYVwSqXtaYbuALVQTJSBaUU5
# DV9gdv/aK0f8k5TkeFr+S598G2l84JB9MYIB2zCCAdcCAQEwKTAVMRMwEQYDVQQD
# DApTZWFuIE1hY2V5AhA7Wkn363I8sU+TMwSw335wMA0GCWCGSAFlAwQCAQUAoIGE
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkE
# MSIEII2WWCec2/30bfcvLlw8bWoMzmyTtTpZorJzGAqFAmOwMA0GCSqGSIb3DQEB
# AQUABIIBAEZXP/BpK+Mw7BiU5SLSqO51naQNhQS4Adh38EKw7TZmXAQMgPq7AHox
# qtMfziDaEQkNhO5D3pd498L08dxUcd14To3Qyw5sOWZAHo3a7uDEfOsWWtNhAVL4
# DSamLeEcu0jPXhmRUV8yUBINeRcFk43h++1MKFore/I2HpfpB6zHR81zrsXwXURV
# 5drFlhJ9+2vBnzcmWY2oQdF/eroxn14e6oudzKHPxibBuWUCuEwgQ8fqv9HmKwII
# Fdlv705VPa3I0TjjW/9FLBuSGFb/rXVbar+jevsP0alLNJfjNFoLe7QRVCjLBtYA
# cvzl+xbEn3gCEJFq5uNcjiJNYQitMoM=
# SIG # End signature block
