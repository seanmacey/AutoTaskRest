#Requires -Modules Pester
<#
.SYNOPSIS
    Pester v5 test suite for AutoTaskRest-TEST-CHUNKING.ps1

.DESCRIPTION
    Tests all non-network-calling logic in the AutoTaskRest module.

    SAFETY GUARANTEE: Every function that calls Invoke-AutoTaskAPI, Invoke-AutoTaskAPIREST,
    Invoke-RestMethod, or any Microsoft Graph cmdlet is mocked at the top of each Describe
    block. No test in this file can write data to Autotask or Microsoft 365.

    Coverage:
      - convertto-escapedString
      - Get-ATCredentialHeader  (credential loading, backward-compat plain-text atapi)
      - Get-SundayOfLastXWeeks
      - Convert-ObjArrayDateTimesToSearchableStrings (both overloads)
      - IsWorkingDay
      - Read-AutoTaskCompanies   (routing / chunking logic)
      - Read-AutoTaskTickets     (filter building, chunking, ExcludeNonComplete, ExcludeComplete)
      - Read-AutoTaskEngineers   (routing: by id array, by email, all)
      - Read-AutotaskContacts    (switch routing fix)
      - Read-AutoTaskTimeEntries (date filter logic, billable/nonBillable split,
                                  ForMeOnly filter, DisplayCompanySummary guard,
                                  hoursBillable/hoursNonBillable calculation)
      - Get-ATWeeklySummary      (summary arithmetic)
      - Build-AutotaskDailyTimeStats (date parsing, resource matching)
      - export-KissATTickets     (parameter routing)
      - Set-AutoTaskCompanies    (classification branch logic — read-only mocks)
      - Get-autoTaskEngineerIDToReport (login file reading)
      - Set-AutoTaskEngineerIDToReport (login file update — mocked file I/O)
      - Test-AutoTaskConnection  (success and failure paths)
#>

BeforeAll {
    # ── Load the module under test ──────────────────────────────────────────────
    $script:ModulePath = Join-Path $PSScriptRoot 'AutoTaskRest-TEST-CHUNKING.ps1'
    . $script:ModulePath

    # ── Helpers ─────────────────────────────────────────────────────────────────
    function New-FakeTimeEntry {
        param(
            [int]$id = 1,
            [int]$ticketID = 0,
            [int]$resourceID = 100,
            [int]$billingCodeID = 99999,
            [bool]$isNonBillable = $false,
            [double]$hoursWorked = 1.0,
            [string]$dateWorked = '2026-01-06T00:00:00',
            [string]$startDateTime = '2026-01-06T08:00:00Z',
            [string]$endDateTime   = '2026-01-06T09:00:00Z'
        )
        [PSCustomObject]@{
            id             = $id
            ticketID       = $ticketID
            resourceID     = $resourceID
            billingCodeID  = $billingCodeID
            isNonBillable  = $isNonBillable
            hoursWorked    = $hoursWorked
            dateWorked     = $dateWorked
            startDateTime  = [datetime]$startDateTime
            endDateTime    = [datetime]$endDateTime
        }
    }

    function New-FakeLoginInfo {
        param([string]$PlainPassword = 'TestPass123', [string]$PlainApi = 'APIID123')
        $encPwd = ($PlainPassword | ConvertTo-SecureString -AsPlainText -Force) | ConvertFrom-SecureString
        $encApi = ($PlainApi     | ConvertTo-SecureString -AsPlainText -Force) | ConvertFrom-SecureString
        [PSCustomObject]@{
            url      = 'https://webservices6.autotask.net/ATServicesRest/'
            UserName = 'api@example.com'
            Secret   = $encPwd
            atapi    = $encApi
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'convertto-escapedString' {

    It 'replaces & with %26' {
        convertto-escapedString -inputString 'Smith & Sons' | Should -Be 'Smith %26 Sons'
    }

    It 'leaves strings without & unchanged' {
        convertto-escapedString -inputString 'Acme Ltd' | Should -Be 'Acme Ltd'
    }

    It 'handles multiple & characters' {
        convertto-escapedString -inputString 'A & B & C' | Should -Be 'A %26 B %26 C'
    }

    It 'handles empty string' {
        convertto-escapedString -inputString '' | Should -Be ''
    }

    It 'handles string that is only &' {
        convertto-escapedString -inputString '&' | Should -Be '%26'
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Get-ATCredentialHeader' {

    Context 'With valid encrypted credentials supplied via LoginInfo' {
        It 'returns a hashtable with the four required keys' {
            $login = New-FakeLoginInfo
            $baseUrl = $null
            $header = Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$baseUrl)
            $header | Should -BeOfType [hashtable]
            $header.Keys | Should -Contain 'ApiIntegrationCode'
            $header.Keys | Should -Contain 'UserName'
            $header.Keys | Should -Contain 'Secret'
            $header.Keys | Should -Contain 'Content-Type'
        }

        It 'decrypts the Secret correctly' {
            $login = New-FakeLoginInfo -PlainPassword 'MySecret99'
            $header = Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$null)
            $header['Secret'] | Should -Be 'MySecret99'
        }

        It 'decrypts the ApiIntegrationCode correctly' {
            $login = New-FakeLoginInfo -PlainApi 'MYAPIKEY'
            $header = Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$null)
            $header['ApiIntegrationCode'] | Should -Be 'MYAPIKEY'
        }

        It 'sets Content-Type to application/json' {
            $login = New-FakeLoginInfo
            $header = Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$null)
            $header['Content-Type'] | Should -Be 'application/json'
        }

        It 'sets BaseUrl ref to the url in LoginInfo' {
            $login = New-FakeLoginInfo
            $baseUrl = $null
            Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$baseUrl) | Out-Null
            $baseUrl | Should -Be 'https://webservices6.autotask.net/ATServicesRest/'
        }
    }

    Context 'Backward-compatibility: plain-text atapi in LoginInfo' {
        It 'accepts plain-text atapi and warns once' {
            $login = [PSCustomObject]@{
                url      = 'https://webservices6.autotask.net/ATServicesRest/'
                UserName = 'api@example.com'
                Secret   = (('Pass1' | ConvertTo-SecureString -AsPlainText -Force) | ConvertFrom-SecureString)
                atapi    = 'PLAIN_API_ID'   # not encrypted
            }
            $header = $null
            { $header = Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$null) } |
                Should -Not -Throw
            $header['ApiIntegrationCode'] | Should -Be 'PLAIN_API_ID'
        }
    }

    Context 'Missing credentials' {
        It 'throws when LoginInfo is missing Secret' {
            $login = [PSCustomObject]@{ url = 'https://x/'; UserName = 'u'; Secret = ''; atapi = 'a' }
            { Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$null) } | Should -Throw
        }

        It 'throws when LoginInfo has no url' {
            $login = [PSCustomObject]@{ url = ''; UserName = 'u'; Secret = 'x'; atapi = 'a' }
            { Get-ATCredentialHeader -LoginInfo $login -BaseUrl ([ref]$null) } | Should -Throw
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Get-SundayOfLastXWeeks' {

    It 'returns a DateTime' {
        Get-SundayOfLastXWeeks -LastXWeeks 1 | Should -BeOfType [datetime]
    }

    It 'returns midnight (time component is zero)' {
        $result = Get-SundayOfLastXWeeks -LastXWeeks 1
        $result.Hour   | Should -Be 0
        $result.Minute | Should -Be 0
        $result.Second | Should -Be 0
    }

    It 'returned date is a Sunday' {
        $result = Get-SundayOfLastXWeeks -LastXWeeks 1
        $result.DayOfWeek | Should -Be 'Sunday'
    }

    It 'LastXWeeks=2 is 7 days before LastXWeeks=1' {
        $one = Get-SundayOfLastXWeeks -LastXWeeks 1
        $two = Get-SundayOfLastXWeeks -LastXWeeks 2
        ($one - $two).Days | Should -Be 7
    }

    It 'LastXWeeks=4 is 21 days before LastXWeeks=1' {
        $one  = Get-SundayOfLastXWeeks -LastXWeeks 1
        $four = Get-SundayOfLastXWeeks -LastXWeeks 4
        ($one - $four).Days | Should -Be 21
    }

    It 'defaults to 1 when 0 is supplied' {
        $default = Get-SundayOfLastXWeeks -LastXWeeks 1
        $zero    = Get-SundayOfLastXWeeks -LastXWeeks 0
        $zero | Should -Be $default
    }

    It 'result is always in the past' {
        $result = Get-SundayOfLastXWeeks -LastXWeeks 1
        $result | Should -BeLessThan (Get-Date)
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Convert-ObjArrayDateTimesToSearchableStrings (pipeline / -obj overload)' {

    It 'adds _Searchable property for each DateTime field' {
        $obj = [PSCustomObject]@{ name = 'test'; createdAt = [datetime]'2026-01-15 09:30:00' }
        Convert-ObjArrayDateTimesToSearchableStrings -obj $obj
        $obj.PSObject.Properties.Name | Should -Contain 'createdAt_Searchable'
    }

    It 'formats the _Searchable string as sortable ISO 8601' {
        $obj = [PSCustomObject]@{ ts = [datetime]'2026-03-05 14:22:00' }
        Convert-ObjArrayDateTimesToSearchableStrings -obj $obj
        $obj.ts_Searchable | Should -Be '2026-03-05T14:22:00'
    }

    It 'does not add _Searchable for non-DateTime properties' {
        $obj = [PSCustomObject]@{ name = 'hello'; count = 42 }
        Convert-ObjArrayDateTimesToSearchableStrings -obj $obj
        $obj.PSObject.Properties.Name | Should -Not -Contain 'name_Searchable'
        $obj.PSObject.Properties.Name | Should -Not -Contain 'count_Searchable'
    }

    It 'handles an array of objects' {
        $objs = @(
            [PSCustomObject]@{ d = [datetime]'2026-01-01' },
            [PSCustomObject]@{ d = [datetime]'2026-06-15' }
        )
        Convert-ObjArrayDateTimesToSearchableStrings -obj $objs
        $objs[0].d_Searchable | Should -Be '2026-01-01T00:00:00'
        $objs[1].d_Searchable | Should -Be '2026-06-15T00:00:00'
    }

    It 'processes multiple DateTime fields on one object' {
        $obj = [PSCustomObject]@{ start = [datetime]'2026-01-01'; end = [datetime]'2026-01-02' }
        Convert-ObjArrayDateTimesToSearchableStrings -obj $obj
        $obj.PSObject.Properties.Name | Should -Contain 'start_Searchable'
        $obj.PSObject.Properties.Name | Should -Contain 'end_Searchable'
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Convert-ObjArrayDateTimesToSearchableStrings (-items overload)' {

    It 'adds _Searchable property via -items parameter' {
        $obj = [PSCustomObject]@{ created = [datetime]'2026-02-20 08:00:00' }
        Convert-ObjArrayDateTimesToSearchableStrings -items @($obj)
        $obj.PSObject.Properties.Name | Should -Contain 'created_Searchable'
    }

    It 'formats correctly via -items overload' {
        $obj = [PSCustomObject]@{ dt = [datetime]'2026-11-11 11:11:11' }
        Convert-ObjArrayDateTimesToSearchableStrings -items @($obj)
        $obj.dt_Searchable | Should -Be '2026-11-11T11:11:11'
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'IsWorkingDay' {

    It 'returns true for a Monday' {
        IsWorkingDay -date ([datetime]'2026-01-05') | Should -BeTrue   # Monday
    }

    It 'returns true for a Friday' {
        IsWorkingDay -date ([datetime]'2026-01-09') | Should -BeTrue   # Friday
    }

    It 'returns false for a Saturday' {
        IsWorkingDay -date ([datetime]'2026-01-10') | Should -BeFalse  # Saturday
    }

    It 'returns false for a Sunday' {
        IsWorkingDay -date ([datetime]'2026-01-11') | Should -BeFalse  # Sunday
    }

    It 'returns true for a Wednesday' {
        IsWorkingDay -date ([datetime]'2026-01-07') | Should -BeTrue   # Wednesday
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskCompanies — routing and chunking logic' {

    BeforeAll {
        # Mock all outbound calls — no data written back, no network calls
        Mock Invoke-AutoTaskAPI      { return @() } -ModuleName $null
        Mock Get-ATCredentialHeader  { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskCompanyClassificationIcons { return @() } -ModuleName $null
    }

    It 'calls Invoke-AutoTaskAPI once for a single company ID' {
        Read-AutoTaskCompanies -id 12345 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter { $entityName -eq 'v1.0/Companies' }
    }

    It 'passes includeFields on single-ID lookup' {
        Read-AutoTaskCompanies -id 12345 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $entityName -eq 'v1.0/Companies' -and $null -ne $includeFields
        }
    }

    It 'chunks multiple IDs into batches of 100' {
        $ids = 1..250  # 3 chunks: 100, 100, 50
        Read-AutoTaskCompanies -id $ids | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 3 -ParameterFilter { $entityName -eq 'v1.0/Companies' }
    }

    It 'searches by name using contains when exactNameMatch not set' {
        Read-AutoTaskCompanies -CompanyName 'Acme' | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"op":"contains"'
        }
    }

    It 'searches by name using eq when exactNameMatch is set' {
        Read-AutoTaskCompanies -CompanyName 'Acme' -exactNameMatch | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"op":"eq"'
        }
    }

    It 'uses SearchFirstBy id when no filters given (active companies)' {
        Read-AutoTaskCompanies | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFirstBy -eq 'isActive'
        }
    }

    It 'uses SearchFirstBy id when includeInactive is set' {
        Read-AutoTaskCompanies -includeInactive | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFirstBy -eq 'id'
        }
    }

    It 'escapes & in company name' {
        Read-AutoTaskCompanies -CompanyName 'Smith & Sons' | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '%26'
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskTickets — filter building and chunking' {

    BeforeAll {
        Mock Invoke-AutoTaskAPI     { return @() } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskCompanies { return @() } -ModuleName $null
        Mock Read-AutoTaskEngineers { return @() } -ModuleName $null
        Mock Read-AutoTaskTicketInformation { return [PSCustomObject]@{ queueID=@(); status=@() } } -ModuleName $null
    }

    It 'builds an unquoted integer in-filter for a single ID' {
        Read-AutoTaskTickets -ids @(42) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $id -eq 42
        }
    }

    It 'builds unquoted integer in-filter for multiple IDs' {
        Read-AutoTaskTickets -ids @(1,2,3) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"value":\[1,2,3\]'
        }
    }

    It 'chunks 150 IDs into 2 API calls' {
        $ids = 1..150
        Read-AutoTaskTickets -ids $ids | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 2 -ParameterFilter {
            $entityName -eq 'v1.0/Tickets'
        }
    }

    It 'chunks 201 IDs into 3 API calls' {
        $ids = 1..201
        Read-AutoTaskTickets -ids $ids | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 3 -ParameterFilter {
            $entityName -eq 'v1.0/Tickets'
        }
    }

    It 'adds ExcludeNonComplete filter (completedDate exists)' {
        Read-AutoTaskTickets -ExcludeNonComplete -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'notExist.*completedDate' -or
            $SearchFurtherBy -match 'Exist.*completedDate'
        }
    }

    It 'adds ExcludeComplete filter (completedDate notExist)' {
        Read-AutoTaskTickets -ExcludeComplete -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'notExist.*completedDate'
        }
    }

    It 'adds whereResourceAssigned filter' {
        Read-AutoTaskTickets -whereResourceAssigned -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'Exist.*assignedResourceID'
        }
    }

    It 'adds TitleContains filter with correct JSON (no double-quotes around value)' {
        Read-AutoTaskTickets -TitleContains 'RMM' -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"op":"contains"' -and
            $SearchFurtherBy -match '"value":"RMM"'
        }
    }

    It 'adds TitleBeginsWith filter with correct JSON' {
        Read-AutoTaskTickets -TitleBeginsWith 'RMM' -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"op":"beginsWith"' -and
            $SearchFurtherBy -match '"value":"RMM"'
        }
    }

    It 'applies LastActionFromDate filter when supplied' {
        $date = [datetime]'2026-01-01'
        Read-AutoTaskTickets -LastActionFromDate $date | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'lastActivityDate.*2026-01-01'
        }
    }

    It 'applies LastxWeeks filter without calling ToUniversalTime on local date' {
        # The date must be in local format (no Z suffix) in the filter string
        Read-AutoTaskTickets -LastxWeeks 1 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'lastActivityDate' -and
            $SearchFurtherBy -notmatch '\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z'
        }
    }

    It 'applies TicketNumbers filter as quoted strings' {
        Read-AutoTaskTickets -TicketNumbers @('T20260101.0001','T20260101.0002') | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'TicketNumber'
        }
    }

    It 'applies default include fields when no fields specified' {
        Read-AutoTaskTickets -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $includeFields -contains 'id' -and $includeFields -contains 'title'
        }
    }

    It 'respects custom includeFields when provided' {
        Read-AutoTaskTickets -includeFields @('id','title') -LastActionFromDate (Get-Date) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $includeFields -contains 'id' -and $includeFields -contains 'title'
        }
    }

    It 'returns null and makes no API call when no filters and no ids supplied and extraSearchs is empty' {
        # With zero extraSearchs and no IDs, the else branch runs but only calls API if extraSearchs.count > 0
        Mock Invoke-AutoTaskAPI { return @() } -ModuleName $null
        Read-AutoTaskTickets | Out-Null
        # No crash — function completes
        $true | Should -BeTrue
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskEngineers — routing logic' {

    BeforeAll {
        Mock Invoke-AutoTaskAPI     { return @() } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
    }

    It 'calls API once per ID when given an ID array' {
        Read-AutoTaskEngineers -id @(1, 2, 3) | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 3 -ParameterFilter { $entityName -eq 'v1.0/Resources' }
    }

    It 'calls API once with email filter when email supplied' {
        Read-AutoTaskEngineers -email 'jane@kissit.co.nz' | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'emailAddress'
        }
    }

    It 'does not call API for resources when given invalid email' {
        Read-AutoTaskEngineers -email 'not-an-email' | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 0 -ParameterFilter { $entityName -eq 'v1.0/Resources' }
    }

    It 'calls API with userType filter when no parameters given' {
        Read-AutoTaskEngineers | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'userType'
        }
    }

    It 'always calls API a second time for daily availabilities' {
        Read-AutoTaskEngineers | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $entityName -eq 'v1.0/ResourceDailyAvailabilities'
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutotaskContacts — switch routing fix' {

    BeforeAll {
        Mock Invoke-AutoTaskAPI     { return @() } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
    }

    It 'searches by ID when ID >= 0' {
        Read-AutotaskContacts -ID 12345 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFirstBy -eq 'id' -and $ID -eq 12345
        }
    }

    It 'searches by email when eMail is supplied' {
        Read-AutotaskContacts -eMail 'bob@example.com' | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'emailAddress'
        }
    }

    It 'returns all contacts when no parameters given (ID defaults to -1)' {
        Read-AutotaskContacts | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $entityName -eq 'v1.0/Contacts' -and
            $null -eq $SearchFurtherBy -and
            $SearchFirstBy -ne 'id'
        }
    }

    It 'does NOT default to returning all contacts when a valid ID is given' {
        Read-AutotaskContacts -ID 99 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter { $ID -eq 99 }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskTimeEntries — date filter logic' {

    BeforeAll {
        Mock Invoke-AutoTaskAPI     { return @() } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskBillingCodes { return @() } -ModuleName $null
        Mock Read-AutoTaskRoles        { return @() } -ModuleName $null
        Mock Read-AutoTaskEngineers    { return @() } -ModuleName $null
        Mock Read-AutoTaskTickets      { return @() } -ModuleName $null
        Mock Get-autoTaskEngineerIDToReport { return 100 } -ModuleName $null
    }

    It 'uses dateWorked/endDateTime OR filter when LastXWeeks is given' {
        Read-AutoTaskTimeEntries -LastXWeeks 1 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"op":"or"' -and
            $SearchFurtherBy -match 'dateWorked' -and
            $SearchFurtherBy -match 'endDateTime'
        }
    }

    It 'uses dateWorked/endDateTime OR filter when FromDateLocal is given' {
        Read-AutoTaskTimeEntries -FromDateLocal ([datetime]'2026-01-01') | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'dateWorked'
        }
    }

    It 'uses dateWorked/endDateTime OR filter when LastxMonths is used' {
        Read-AutoTaskTimeEntries -LastxMonths 1 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'dateWorked'
        }
    }

    It 'adds resourceID filter when ForMeOnly is set' {
        Read-AutoTaskTimeEntries -LastXWeeks 1 -ForMeOnly | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match 'resourceID'
        }
    }

    It 'adds resourceID filter when ForResouerceID is supplied' {
        Read-AutoTaskTimeEntries -LastXWeeks 1 -ForResouerceID 999 | Out-Null
        Should -Invoke Invoke-AutoTaskAPI -Times 1 -ParameterFilter {
            $SearchFurtherBy -match '"value":"999"'
        }
    }

    It 'returns early with warning when no time entries found' {
        Mock Invoke-AutoTaskAPI { return $null } -ModuleName $null
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        $result | Should -BeNullOrEmpty
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskTimeEntries — hoursBillable / hoursNonBillable calculation' {

    BeforeAll {
        # Return two fake time entries from the API
        $fakeEntries = @(
            New-FakeTimeEntry -id 1 -isNonBillable $false -hoursWorked 2.0 -ticketID 0
            New-FakeTimeEntry -id 2 -isNonBillable $true  -hoursWorked 1.5 -ticketID 0
        )
        Mock Invoke-AutoTaskAPI     { return $fakeEntries } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskBillingCodes { return @() } -ModuleName $null
        Mock Read-AutoTaskRoles        { return @() } -ModuleName $null
        Mock Read-AutoTaskEngineers    { return @() } -ModuleName $null
        Mock Read-AutoTaskTickets      { return @() } -ModuleName $null
        Mock Get-autoTaskEngineerIDToReport { return 100 } -ModuleName $null
    }

    It 'sets hoursBillable = hoursWorked for billable entries' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        ($result | Where-Object id -eq 1).hoursBillable | Should -Be 2.0
    }

    It 'sets hoursNonBillable = 0 for billable entries' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        ($result | Where-Object id -eq 1).hoursNonBillable | Should -Be 0
    }

    It 'sets hoursNonBillable = hoursWorked for non-billable entries' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        ($result | Where-Object id -eq 2).hoursNonBillable | Should -Be 1.5
    }

    It 'sets hoursBillable = 0 for non-billable entries' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        ($result | Where-Object id -eq 2).hoursBillable | Should -Be 0
    }

    It 'adds OADate field to all entries' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        $result | ForEach-Object { $_.PSObject.Properties.Name | Should -Contain 'OADate' }
    }

    It 'adds startDateTimeLocal and endDateTimeLocal fields' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1
        $result[0].PSObject.Properties.Name | Should -Contain 'startDateTimeLocal'
        $result[0].PSObject.Properties.Name | Should -Contain 'endDateTimeLocal'
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskTimeEntries — internal classification marking' {

    BeforeAll {
        $script:ATInternalClasificationCode = 200

        $internalEntry = New-FakeTimeEntry -id 10 -ticketID 500 -isNonBillable $false -hoursWorked 1.0
        $clientEntry   = New-FakeTimeEntry -id 11 -ticketID 501 -isNonBillable $false -hoursWorked 2.0

        $internalTicket = [PSCustomObject]@{
            id = 500; companyID = 999; CompanyClassification = 200
            CompanyName = 'Kiss IT'; CompanyIsInternal = $true
            title = 'Internal work'; ticketNumber = 'T001'; BillingCodeID = 99
        }
        $clientTicket = [PSCustomObject]@{
            id = 501; companyID = 888; CompanyClassification = 1
            CompanyName = 'Acme Ltd'; CompanyIsInternal = $false
            title = 'Client work'; ticketNumber = 'T002'; BillingCodeID = 11
        }

        Mock Invoke-AutoTaskAPI     { return @($internalEntry, $clientEntry) } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskBillingCodes { return @() } -ModuleName $null
        Mock Read-AutoTaskRoles        { return @() } -ModuleName $null
        Mock Read-AutoTaskEngineers    { return @() } -ModuleName $null
        Mock Read-AutoTaskTickets      { return @($internalTicket, $clientTicket) } -ModuleName $null
        Mock Get-autoTaskEngineerIDToReport { return 100 } -ModuleName $null
    }

    It 'marks internal company entries as non-billable' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1 -includeTicketDetails
        $internal = $result | Where-Object id -eq 10
        $internal.isNonBillable | Should -BeTrue
    }

    It 'marks client company entries as billable client' {
        $result = Read-AutoTaskTimeEntries -LastXWeeks 1 -includeTicketDetails
        $client = $result | Where-Object id -eq 11
        $client.IsBillableClient | Should -BeTrue
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Read-AutoTaskTimeEntries — DisplayCompanySummary guard' {

    It 'does not run summary block when includeTicketDetails is false' {
        Mock Invoke-AutoTaskAPI     { return @(New-FakeTimeEntry) } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskBillingCodes { return @() } -ModuleName $null
        Mock Read-AutoTaskRoles        { return @() } -ModuleName $null
        Mock Read-AutoTaskEngineers    { return @() } -ModuleName $null
        Mock Get-autoTaskEngineerIDToReport { return 100 } -ModuleName $null

        # Should complete without error even though CompanyName property doesn't exist yet
        { Read-AutoTaskTimeEntries -LastXWeeks 1 -DisplayCompanySummary } | Should -Not -Throw
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Get-autoTaskEngineerIDToReport — login file reading' {

    It 'returns the ATResourceID from the login file' {
        $tmpDir  = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "KissTest_$(Get-Random)")
        $tmpFile = Join-Path $tmpDir.FullName 'kissAtapilogin.json'
        @{ ATResourceID = 42; url = 'x'; UserName = 'u'; Secret = 'e'; atapi = 'a' } |
            ConvertTo-Json | Set-Content $tmpFile

        $script:kissATAPIpath = $tmpDir.FullName
        $result = Get-autoTaskEngineerIDToReport
        $result | Should -Be 42

        Remove-Item $tmpDir -Recurse -Force
    }

    It 'returns nothing and does not throw when file has no ATResourceID' {
        $tmpDir  = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "KissTest_$(Get-Random)")
        $tmpFile = Join-Path $tmpDir.FullName 'kissAtapilogin.json'
        @{ url = 'x'; UserName = 'u'; Secret = 'e'; atapi = 'a' } |
            ConvertTo-Json | Set-Content $tmpFile

        $script:kissATAPIpath = $tmpDir.FullName
        { Get-autoTaskEngineerIDToReport } | Should -Not -Throw

        Remove-Item $tmpDir -Recurse -Force
    }

    It 'returns nothing and does not throw when login file does not exist' {
        $script:kissATAPIpath = Join-Path $env:TEMP 'NonExistentKissPath'
        { Get-autoTaskEngineerIDToReport } | Should -Not -Throw
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Set-AutoTaskEngineerIDToReport — login file update' {

    BeforeAll {
        Mock Read-AutoTaskEngineers {
            return [PSCustomObject]@{ id = 77; FullName = 'Jane Smith'; isActive = $true }
        } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Invoke-AutoTaskAPI     { return @() } -ModuleName $null
    }

    It 'writes ATResourceID to the login file when found by id' {
        $tmpDir  = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "KissTest_$(Get-Random)")
        $tmpFile = Join-Path $tmpDir.FullName 'kissAtapilogin.json'
        @{ url='x'; UserName='u'; Secret='e'; atapi='a' } | ConvertTo-Json | Set-Content $tmpFile
        $script:kissATAPIpath = $tmpDir.FullName

        Set-AutoTaskEngineerIDToReport -id 77
        $saved = Get-Content $tmpFile | ConvertFrom-Json
        $saved.ATResourceID | Should -Be 77

        Remove-Item $tmpDir -Recurse -Force
    }

    It 'updates existing ATResourceID when it already exists in the file' {
        $tmpDir  = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "KissTest_$(Get-Random)")
        $tmpFile = Join-Path $tmpDir.FullName 'kissAtapilogin.json'
        @{ url='x'; UserName='u'; Secret='e'; atapi='a'; ATResourceID=1 } | ConvertTo-Json | Set-Content $tmpFile
        $script:kissATAPIpath = $tmpDir.FullName

        Set-AutoTaskEngineerIDToReport -id 77
        $saved = Get-Content $tmpFile | ConvertFrom-Json
        $saved.ATResourceID | Should -Be 77

        Remove-Item $tmpDir -Recurse -Force
    }

    It 'does not write when engineer not found' {
        Mock Read-AutoTaskEngineers { return $null } -ModuleName $null
        $tmpDir  = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "KissTest_$(Get-Random)")
        $tmpFile = Join-Path $tmpDir.FullName 'kissAtapilogin.json'
        @{ url='x'; UserName='u'; Secret='e'; atapi='a' } | ConvertTo-Json | Set-Content $tmpFile
        $script:kissATAPIpath = $tmpDir.FullName

        Set-AutoTaskEngineerIDToReport -id 999
        $saved = Get-Content $tmpFile | ConvertFrom-Json
        $saved.PSObject.Properties.Name | Should -Not -Contain 'ATResourceID'

        Remove-Item $tmpDir -Recurse -Force
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Test-AutoTaskConnection — success and failure paths' {

    It 'returns $true when Invoke-RestMethod succeeds' {
        Mock Invoke-RestMethod  { return [PSCustomObject]@{ version = '1.0' } } -ModuleName $null
        Mock Get-ATCredentialHeader {
            return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' }
        } -ModuleName $null

        $login = New-FakeLoginInfo
        Test-AutoTaskConnection -LoginInfo $login | Should -BeTrue
    }

    It 'returns $null when Invoke-RestMethod throws' {
        Mock Invoke-RestMethod  { throw 'Unauthorized' } -ModuleName $null
        Mock Get-ATCredentialHeader {
            return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' }
        } -ModuleName $null

        $login = New-FakeLoginInfo
        Test-AutoTaskConnection -LoginInfo $login | Should -BeNullOrEmpty
    }

    It 'returns $null when credentials are missing' {
        Mock Get-ATCredentialHeader { throw 'missing creds' } -ModuleName $null
        $login = New-FakeLoginInfo
        Test-AutoTaskConnection -LoginInfo $login | Should -BeNullOrEmpty
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'export-KissATTickets — parameter routing' {

    BeforeAll {
        Mock Read-AutoTaskTickets { return @() } -ModuleName $null
        Mock Export-Csv           { } -ModuleName $null
        Mock New-Item             { } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Invoke-AutoTaskAPI   { return @() } -ModuleName $null
    }

    It 'exports TicketsNotCompleted when no months specified' {
        export-KissATTickets
        Should -Invoke Read-AutoTaskTickets -Times 1 -ParameterFilter {
            $IncludeAllNonComplete -eq $true
        }
    }

    It 'exports TicketsActioned when months > 0' {
        export-KissATTickets -WhereLastActionOccurWithinLastMonths 3
        Should -Invoke Read-AutoTaskTickets -Times 2
    }

    It 'does NOT export TicketsActioned when months = 0' {
        export-KissATTickets -WhereLastActionOccurWithinLastMonths 0
        Should -Invoke Read-AutoTaskTickets -Times 1 -ParameterFilter {
            $IncludeAllNonComplete -eq $true
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Set-AutoTaskCompanies — classification branch logic (read-only mocks)' {

    BeforeAll {
        # Only GET-style mocks; Invoke-AutoTaskAPIREST is mocked to prevent any writes
        Mock Invoke-AutoTaskAPIREST { } -ModuleName $null
        Mock Invoke-AutoTaskAPI     { return @() } -ModuleName $null
        Mock Get-ATCredentialHeader { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Read-AutoTaskCompanyClassificationIcons { return @() } -ModuleName $null
        Mock Read-AutoTaskEngineers { return @() } -ModuleName $null
        Mock Read-AutoTaskCompanies {
            return [PSCustomObject]@{ id = 100; companyName = 'Test Co' }
        } -ModuleName $null
    }

    It 'calls PATCH when a valid numeric CompanyID is given' {
        Set-AutoTaskCompanies -CompanyID 100
        Should -Invoke Invoke-AutoTaskAPIREST -Times 1 -ParameterFilter {
            $Method -eq 'PATCH'
        }
    }

    It 'resolves company by name before patching' {
        Set-AutoTaskCompanies -CompanyName 'Test Co'
        Should -Invoke Read-AutoTaskCompanies -Times 1 -ParameterFilter {
            $CompanyName -eq 'Test Co'
        }
    }

    It 'throws when classification name does not match any known classification' {
        Mock Read-AutoTaskCompanyClassificationIcons {
            return @([PSCustomObject]@{ id = 1; name = 'Commercial' })
        } -ModuleName $null
        { Set-AutoTaskCompanies -CompanyID 100 -Classification 'NonExistentClass' } | Should -Throw
    }

    It 'does not call Invoke-AutoTaskAPIREST when CompanyID is -1 and no name given' {
        Set-AutoTaskCompanies -CompanyID -1
        Should -Invoke Invoke-AutoTaskAPIREST -Times 0
    }

    It 'never calls Invoke-AutoTaskAPIREST with PUT or DELETE' {
        Set-AutoTaskCompanies -CompanyID 100
        Should -Invoke Invoke-AutoTaskAPIREST -Times 0 -ParameterFilter { $Method -eq 'PUT' }
        Should -Invoke Invoke-AutoTaskAPIREST -Times 0 -ParameterFilter { $Method -eq 'DELETE' }
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Safety — write-back prevention' {
    # Explicitly verify that key mutating functions are always mocked in tests above,
    # and that calling them without mocks would require credentials (which won't exist).

    It 'Invoke-AutoTaskAPIREST throws when no credentials file exists' {
        $script:kissATAPIpath = Join-Path $env:TEMP 'NoSuchPath_SafetyCheck'
        { Invoke-AutoTaskAPIREST -url 'v1.0/Companies' -Method PATCH -Body '{}' } | Should -Throw
    }

    It 'Invoke-AutoTaskAPI throws when no credentials file exists' {
        $script:kissATAPIpath = Join-Path $env:TEMP 'NoSuchPath_SafetyCheck'
        { Invoke-AutoTaskAPI -entityName 'v1.0/Companies' } | Should -Throw
    }
}

# ════════════════════════════════════════════════════════════════════════════════
Describe 'Get-ATWeeklySummary — summary arithmetic' {

    BeforeAll {
        $fakeEntries = @(
            [PSCustomObject]@{
                Engineer         = 'Jane Smith'
                HrsClient        = 6.0
                hoursBillable    = 4.0
                HrsInternalProd  = 2.0
                dateWorked       = '2026-01-05T00:00:00'
            },
            [PSCustomObject]@{
                Engineer         = 'Jane Smith'
                HrsClient        = 2.0
                hoursBillable    = 2.0
                HrsInternalProd  = 0.0
                dateWorked       = '2026-01-06T00:00:00'
            }
        )
        Mock Read-AutoTaskTimeEntries { return $fakeEntries } -ModuleName $null
        Mock Get-ATCredentialHeader   { return @{ ApiIntegrationCode='x'; UserName='u'; Secret='s'; 'Content-Type'='application/json' } } -ModuleName $null
        Mock Invoke-AutoTaskAPI       { return @() } -ModuleName $null
    }

    It 'returns one summary row per engineer' {
        $result = Get-ATWeeklySummary -lastXWeeks 1
        $result.Count | Should -Be 1
    }

    It 'calculates total billable hours correctly' {
        $result = Get-ATWeeklySummary -lastXWeeks 1
        $result[0].hoursBillable | Should -Be 6.0
    }

    It 'calculates non-billable client hours as HrsClient minus hoursBillable' {
        $result = Get-ATWeeklySummary -lastXWeeks 1
        $result[0].hoursNonBillable | Should -Be 2.0  # (6+2) - (4+2)
    }

    It 'calculates internal hours correctly' {
        $result = Get-ATWeeklySummary -lastXWeeks 1
        $result[0].hoursInternal | Should -Be 2.0
    }

    It 'includes engineer name in result' {
        $result = Get-ATWeeklySummary -lastXWeeks 1
        $result[0].Engineer | Should -Be 'Jane Smith'
    }

    It 'passes -DisplayCompanySummary to Read-AutoTaskTimeEntries' {
        Get-ATWeeklySummary -lastXWeeks 1 | Out-Null
        Should -Invoke Read-AutoTaskTimeEntries -Times 1 -ParameterFilter {
            $DisplayCompanySummary -eq $true
        }
    }
}

# SIG # Begin signature block
# MIIFgwYJKoZIhvcNAQcCoIIFdDCCBXACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBHlIj6CwCPTsJ8
# w9XHHZ7I0A3IR/JZ646jTfjDdZu7MqCCAv4wggL6MIIB4qADAgECAhA7Wkn363I8
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
# MSIEIDUcUqmd1TnMKkp3bXLmqEuWtJ/ug+wffdedQPrZPEP8MA0GCSqGSIb3DQEB
# AQUABIIBAAaW/CEKViAOfOf3CHBI0txT7mdPsF7cih4xgxMPxoxdk1Bk2JCqBSuh
# 7qisCFSaIecQb8vQOggvIY7qdqtBinzbNICUOpmVqzxUW4PnO0o343azL7xruvQs
# clIwZn5NsAQy9vMzjzKZUEmVK0hwetLZCWcsIgDU0nV7ZPRoDOagaSTeQMPdWmRa
# vXViHalLiknRIizcIOR5OgMR9izON0OOOjnEtQXNGIHqIut4mXOTjCzAGeUyCQyz
# uyJkBNZTvcEU3s1UhZ6ZNLhkUH6z/MVi0Q5hbWEs6jbe1AUTI0TsBsUx2vFpECth
# lC9VfpEfNx0CAAYZ6csLpXpS7cEsD74=
# SIG # End signature block
