$global:kissATAPIpath = "$home\kiss-atapi"
$global:kissATAPIfile = 'kissAtapilogin.json'


# check for to test REST API  https://webservices6.autotask.net/ATServicesRest/swagger/ui/index
# Check for REST API information - such as entitis and calling methods and syntax
#https://autotask.net/help/DeveloperHelp/Content/APIs/REST/REST_API_Home.htm
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

function Invoke-AutoTaskAPIREST() {
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
    $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
    if ($jsn) { $saveobj = $jsn | ConvertFrom-Json }
    if ($saveobj.url -and $saveobj.secret -and $saveobj.username -and $saveobj.atapi) {
        #saved data exists and is valid , so import it
        $saveobj.secret = $saveobj.Secret | Convertto-SecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($saveobj.secret))
        $saveobj.Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        #  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR) 
        write-debug "Invoke-AutoTaskAPIPut: url is $($saveobj.url)"

        $kissATheader = @{'ApiIntegrationCode' = $saveobj.atapi #| Convertto-SecureString | ConvertFrom-SecureString -AsPlainText
            'UserName'                         = $saveobj.UserName
            'Secret'                           = $saveobj.Secret #| Convertto-SecureString #| ConvertFrom-SecureString #-AsPlainText
            'Content-Type'                     = "application/json"
        }

    }
    $url2 = "$($saveobj.url)$Url"
    write-verbose "Invoke-AutoTaskAPIREST $Method $url2 `r`n BODY $body"
    $result = Invoke-RestMethod -Method $Method -Uri $url2  -Headers $kissATheader -Body $Body
    write-verbose "Invoke-AutoTaskAPIREST resultitem = $($result.itemid)"
    $result

}

function Invoke-AutoTaskAPI {
    <#
    .SYNOPSIS
    connects to the Autotask Servers, and 
    
    .DESCRIPTION
    Long description
    Recursively calls autotask for ALL data since Autotask API only releases 500 record at a time

    because Powrshell does not handle processing date fields well => extracting to CSV or parsing a date into a function may be problematic due to date local and format issues.
    if dates are present in the data: then use function Convert-ObjArrayDateTimesToSearchableStrings to convert any of theose date entries to a searchable text string (avoid using Datetime object)
    
    .PARAMETER url
    is supplied then this is the only parameter needed in order to get data from autotask
    this is used by  recursive calls to this subrutine - since every 
    
    .PARAMETER urlStart
    the beginning part of the url to autotask
    this is different for every global region - but constand within
    for NZ it is https://webservices6.autotask.net/atservicesrest/v1.0/

    the value for the region and User  you are running this http Get from is available by calling 
    there is no authorisation needed in order to call that http Get
    example for user gokypolmtounjb6@KISSIT.CO.NZ http://webservices.autotask.net/atservicesrest/v1.0/zoneInformation?user=gokypolmtounjb6@KISSIT.CO.NZ
    
    .PARAMETER entityName
    the name of the autotask entity you are wanting data on
    
    .PARAMETER ID
    A specific ID : athis is only used when you want to find a specific record
    
    .PARAMETER isActive
    is true will only return active records
    
    .PARAMETER SearchFirstBy
    a selectable default to search all by id sequence, or to search only for active records, or to search for nothing (if NOther is used then you must add customised SearchFurtherby paramater)
    
    .PARAMETER SearchFurtherBy
    allows customistaion of serach paramater
    
    .PARAMETER includeFields
    allows the amount of fields returned to be limited, if not used then all fields are returned
    
    .PARAMETER LoopCount
    specifies how many times this function will be recursively called before giving up: defasult is 40
    thesre is a limit like this so as to avoid for ever loops....

    .PARAMETER LoginINfo
    OPTIONAL : An object containing the header details needed for authentication - if this is not used then the saved details will be
    
    .EXAMPLE
    Invoke-AutoTaskAPI -entityName ClassificationIcons -includeFields "id", "name"
    Invoke-AutoTaskAPI -entityName "Companies"  -id $id 
    Invoke-AutoTaskAPI -entityName "Companies"  -includeFields $includeFields -SearchFirstBy Nothing  -SearchFurtherBy "{""op"":""$op"",""Field"":""companyName"",""value"":""$companyName""}"
    $rc = Invoke-AutoTaskAPI -entityName "Companies"  -includeFields $includeFields -SearchFirstBy id 
    Invoke-AutoTaskAPI -entityName "Companies"  -includeFields $includeFields -SearchFirstBy isActive 

    .NOTES
    General notes
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
        [Int32]
        $ID = -1,

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

    if ($LoginINfo) {
        Write-Verbose "Invoke-AutoTask: using the paramatised LoginInfo instead of the pre-saved authentication details"
        $saveobj = $LoginINfo
    }
    elseif (test-path -path "$kissATAPIpath\$kissATAPIfile" ) {
        $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
        if ($jsn) { $saveobj = $jsn | ConvertFrom-Json }
        if ($saveobj.url -and $saveobj.secret -and $saveobj.username -and $saveobj.atapi) {
            #saved data exists and is valid , so import it
            # $saveobj = $r
            $saveobj.secret = $saveobj.Secret | Convertto-SecureString
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($saveobj.secret))
            $saveobj.Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            #  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR) 
            write-debug "Invoke-AutoTask: secret is $($saveobj.Secret)"
            write-debug "Invoke-AutoTask: userName is $($saveobj.username)"
            write-debug "Invoke-AutoTask: ApiIntegrationCode is $($saveobj.atapi)"
 
        }
    }
    else {
        write-host " **** Invoke-AutoTask: there were no saved credentials"
        Write-Warning "You must first Set-LoginAtotask and save your APID and credentials"
        throw "Invoke-AutoTask: You must first Set-LoginAtotask and save your APID and credentials"
        return
    }

    if (!($saveobj.url) -or !($saveobj.secret) -or !($saveobj.username) -or !($saveobj.atapi)) {
        write-Host "Invoke-AutoTask: Cant login with $($saveobj.username )"
        throw "Invoke-AutoTaskAPI: At least one of the Authentication parameters is missing"
        return
    }

    





 
    # $Plainpassword = $r.Secret | Convertto-SecureString  
    # $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Plainpassword)
    # $Plainpassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    $kissATheader = @{'ApiIntegrationCode' = $saveobj.atapi #| Convertto-SecureString | ConvertFrom-SecureString -AsPlainText
        'UserName'                         = $saveobj.UserName
        'Secret'                           = $saveobj.Secret #| Convertto-SecureString #| ConvertFrom-SecureString #-AsPlainText
        'Content-Type'                     = "application/json"
    }
    # Write-verbose "secret is $($Saveobj.Secret)"


    #----- Delete line once debugging finished
    #  Write-Debug "KissATheader = $($kissATheader |ConvertTo-Json)"
    #----------

    if ($url -and ($returnRaw -eq $true)) {
        Write-Verbose "Invoke-AutoTaskAPI get RAw data based on $url"
        Invoke-RestMethod -Method Get -Uri $url  -Headers $kissATheader  #-SkipHeaderValidation
        Write-Debug "url: $url <br> headers: $kissATheader"
        return
    }
    if ($urlFixedSuffix) {
        $url2 = "$($saveobj.url)$UrlFixedSuffix"
        Write-Verbose "Invoke-AutoTaskAPI get Raw data based on $url2"
        Invoke-RestMethod -Method Get -Uri $url2  -Headers $kissATheader    #-SkipHeaderValidation 
        return


    }
   
    if (($id -gt -1) -and $entityName) { 
        # just return a SINGLE item with a specific ID
        # $url2 = "$urlstart$entityName/$ID"
        $url2 = "$($saveobj.url)$entityName/$ID"
        Write-Verbose "Invoke-AutoTaskAPI getiing just one $entityname item $id : $url2"
        $Result = Invoke-RestMethod -Method Get -Uri $url2  -Headers $kissATheader  #-SkipHeaderValidation #-FollowRelLink
        Write-Debug "url: $url2 <br> headers: $kissATheader"

        Write-Verbose "Invoke-AutoTaskAPI item count=$($result.item.count)"
        if ($ReturnRaw -eq $true) {
            write-host "Invoke-AutoTaskAPI Returning raw data, and not an object collection - this WILL include userDefinedFields"
            return $result
        }
        return $Result.Item
    }
 
    if ($entityName) {
        # prepare a collection of items to return - and might need to be called recursively
        $entityFilter = ''
        switch ($SearchFirstBy) {
            "isActive" {
                Write-Verbose "Invoke-AutoTaskAPI : returning only $entityname items where field isActive = true"
                $entityfilter = '{"op":"eq","field":"isActive","value":"true"}'
            }
            "id" {
                Write-Verbose "Invoke-AutoTaskAPI : returning  $entityname where ID GTE 0 and isactive:$isactive"
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
        $url2 = "$($saveobj.url)$entityName/query?search={$_search}"
        #$url2 = "$urlstart$entityName/query?search={$_search}"
    }
    else { $url2 = $url }
    
    Write-verbose "getting  $entityname items  $url2"
    $Result = Invoke-RestMethod -Method Get -Uri $url2  -Headers $kissATheader  #-SkipHeaderValidation
    $RecordsRecieved = $Result.pageDetails.Count
    $apidata = $Result.items 
    Write-Verbose "retrieved $RecordsRecieved records: which should equal $($apidata.count)"
    Write-Verbose "returned PageDetails `n$($Result.pageDetails |ConvertTo-Json)"

    
    #now prepare the next 500 items
    $Nextpage = $Result.pageDetails.nextPageUrl
    if (($LoopCount -gt 1) -and $Nextpage) {
        Write-Verbose "Invoke-AutoTaskAPI LoopCount Value = $Loopcount"

        if ($CheckDuplicatesOf) {
            $alreadyCapturedData += $apidata
            $apidata += Invoke-AutoTaskAPI -url $Nextpage -LoopCount ($LoopCount - 1) -CheckDuplicatesOf $CheckDuplicatesOf -alreadyCapturedData $alreadyCapturedData
        }
        else {
            $apidata += Invoke-AutoTaskAPI -url $Nextpage -LoopCount ($LoopCount - 1) 
        }


    }
    Write-Verbose "Invoke-AutoTaskAPI total Returned items = $($apidata.count)"
    return $apidata
}


function read-AutoTaskLastQuickNote() {
    [CmdletBinding()]
    param (  )
    $rc = Invoke-AutoTaskAPI -url "https://webservices6.autotask.net/ATServicesRest/V1.0/CompanyNotes/query?search={""IncludeFields"":[""id"",""CompanyID"",""lastModifiedDate"",""Note""],""filter"":[{""op"":""and"",""items"":[{""op"":""eq"",""field"":""actionType"",""value"":""5""},{""op"":""gt"",""field"":""lastModifiedDate"",""value"":""2018-01-01T00:00:00.00Z""}]}]}"
    $rc = $rc | Group-Object companyID

    foreach ($item in $rc) {
        $item[0].Group | Add-Member -NotePropertyName NoteCount -NotePropertyValue $item.Count
        ($item[0].Group | Sort-Object -Descending lastModifiedDate)[0]
    }

}


function Convert-ObjArrayDateTimesToSearchableStrings () {
    <#
    .SYNOPSIS
    change date fields to Seartchable date string fields
        
    .DESCRIPTION
     will work on an array of objects, and checkes every object individually

    
    .PARAMETER obj
    an object array (can also be a single obkect)
    
    .EXAMPLE
    Convert-ObjArrayDateTimesToSearchableStrings -obj $timeentries
    
    .NOTES
    General notes
    #>
    param (
        # Parameter help description
        [Parameter(Mandatory = $true)]
        [psobject[]]
        $obj
    )
    
    ## locate any datetime objects and change them to searchable date strings
    foreach ($item in $obj) {
        $dtfixs = $item | Get-Member -MemberType properties | where-object definition -like "datetime*"
        foreach ($dtfix in $dtfixs) {
            if ($item.$($dtfix.name)) {
                $item.$($dtfix.name) = [string]$item.$($dtfix.name).ToString('s')  #$i.dateWorked.ToString('s')
                # write-host "change format of $item.$($dtfix.name) "
            }   
        }
    }
    
    # return $Obj
}


function Read-AutoTaskCompanyClassificationIcons() {
    [CmdletBinding()]
    param (  )
    $rc = Invoke-AutoTaskAPI -entityName 'v1.0/ClassificationIcons'   -SearchFirstBy id
    $rc
}

function Read-AutoTaskMostRecentCompanyTicket() {
    [CmdletBinding()]
    param (  )
    $rc = Read-AutoTaskTickets -LastActionFromDate "2020-01-01T00:00:00" -Verbose -DontexpandticketInformation -whereResourceAssigned -includeFields ("companyID", "completedDate", "id", "title", "createDate") -DontincludeNonComplete | Group-Object companyID
    #$rc.psobject.properties.remove('userDefinedFields')
    foreach ($item in $rc) {
        ($item[0].Group | Sort-Object -Descending completedDate)[0]
    }
}


function Read-AutoTaskCompanies {
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
    Read-AutoTaskCompanies

     Read-AutoTaskCompanies -CompanyName "imatec" -debug 
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
    [CmdletBinding()]
    param (
        [Parameter()]
        [int]
        $id = -1,
        # Parameter help description
        #[Parameter(AttributeValues)]
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
        { $id -ge 0 } {
            write-verbose "Read-AUtoTaskCompanies - for a single ID $id"
            $rc = Invoke-AutoTaskAPI -entityName 'v1.0/Companies'  -id $id #; break 
            break
        }
        { $CompanyName } {
            write-verbose "Read-AUtoTaskCompanies - for a exact match :$companyName"
            [string]$srch = "{""op"":""$op"",""Field"":""companyName"",""value"":""$companyName""}"  #{"op":"contains","Field":"companyName","value":"imatec"}
            $rc = Invoke-AutoTaskAPI -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy Nothing  -SearchFurtherBy $srch
            break 
        }
        { $includeInactive -eq $true } { 
            write-verbose "Read-AUtoTaskCompanies - for ALL companies including inactive"
            $rc = Invoke-AutoTaskAPI -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy id -CheckDuplicatesOf "id"
            break 
        }
        default {
            write-verbose "Read-AutoTaskCompanies - for ALL Active companies"
            $rc = Invoke-AutoTaskAPI -entityName 'v1.0/Companies'  -includeFields $includeFields -SearchFirstBy isActive -CheckDuplicatesOf "id"
        }
    }

    if ($rc) {
        $branch = ($rc.userDefinedFields | Where-Object { $_.name -eq "Branch" })[0]
        $rc = $rc | select-Object -Property * , @{name = "Branch"; e = { $branch.value } } -ErrorAction SilentlyContinue | Select-Object -ExcludeProperty userDefinedFields
        if (!($DontExpandChildIDFields -eq $true)) {

            
            Convert-ObjArrayDateTimesToSearchableStrings -obj $rc #|Out-Null

            #$rc.userDefinedFields
            #$rc = $rc | Select-Object -ExcludeProperty userDefinedFields

            if ($GetEngineers) {
                $rc | Add-Member -NotePropertyName Primary -NotePropertyValue ""
                $rc | Add-Member -NotePropertyName Secondary -NotePropertyValue ""
                $AllPrimeTechnicians = Read-AutotaskPrimaryEngineers
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
            $classificationIcons = Read-AutoTaskCompanyClassificationIcons
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
        write-verbose "Done Read-AutoTaskCompanies" #-foregroundColor Green

        return $rc
    }
}

function Build-AutoTaskInternalTicketsTime() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [psobject]
        $timeEntries
        #must fields : ticketID, CompanyID, dateworked
        
    )
    # this function modifies rthe $timeEntries obj, it does not need to return it! since the input object is reference, not a value
    if (!($timeEntries.TicketID -and $timeEntries.CompanyID -and $timeEntries.dateWorked)) {
        throw "the timeentires input object is missing either TicketID, CompanyID or dateWorked"
        return $timeEntries
    }

    #insert the username of eachtech
    $Resources = Read-AutoTaskEngineers #| Where-Object { ($_.id -in $TimeEntries.resourceID) }
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
    $CompanyTickets = Read-AutoTaskTickets -LastActionFromDate $earliestDate -CompanyIDs (29762985 , 0, 1, 29740186 , 29761818, 29762138, 29718567, 29762986)
   
    $InternalEntries = $timeEntries | Where-Object TicketID -in $CompanyTickets.id

    foreach ($i in $InternalEntries) {
        $items = $i | Where-Object { (($_.isNonBillable -eq $true) -or ($_.billingCodeID -in $nonBillableCodes)) }
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


    <#
    $groupedtickets = $timeEntries | Group-Object ticketID


    foreach ($item in $groupedtickets) {

        if ($item.name -in $CompanyTickets.id) {
            Write-Verbose "Build-AutoTaskInternalTicketsTime found time entries on an Internal Ticket $($item.name)"
            #    $item
        
            # mark non billable internal tickets
            $subitems = $item.group | Where-Object isNonBillable -eq $true
            foreach ($i in $subitems) {
                #  $i.InternalTicket = $i.hoursWorked
                Write-Debug "Build-AutoTaskInternalTicketsTime  nonBillable Ticket Hrs $($i.hoursWorked)"
                $timeEntries | Add-Member -NotePropertyName 'InternalTicketNobnBillable' -NotePropertyValue $i.hoursWorked -Force

            }
            #mark billable internal tickets
            $subitems = $item.group | Where-Object isNonBillable -eq $false
            foreach ($i in $subitems) {
                #    $i.InternalTicket = $i.hoursWorked
                Write-Debug "Build-AutoTaskInternalTicketsTime billable Ticket Hrs $($i.hoursWorked)"
                $timeEntries | Add-Member -NotePropertyName 'InternalTicketBillable' -NotePropertyValue $i.hoursWorked -Force

            }
        }
    }
    #>
    #  NO Need to return a Value since the input object is alrerady modified (it is a reference object)
    return $timeEntries
}


function Build-AutoTaskRMMTime() {
    <#
    .SYNOPSIS
    updates timeentries field RMM with hoursworked, if and only if the title of the ticket starts with RMM
    
    
    .DESCRIPTION
    updates timeentries field RMM with hoursworked, if and only if the title of the ticket starts with RMM
    it does not update the RMM timeentries column is the RMM text is only conatined in the title field of the ticket - thus to record hours against RMM activity the ticket title must start with RMM
 
    
    .PARAMETER timeEntries
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    # gets any tickets where the Title starts with RMM
    #does not return any ti
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [psobject]
        $timeEntries,
        [int[]]
        $RMMtaskCodes = 29712660
    )
    $earliestDate = ($timeEntries | Measure-Object dateWorked -min).Minimum
    $timeEntries | Add-Member -NotePropertyName 'RMMTicket' -NotePropertyValue 0.0 -Force
    $timeEntries | Add-Member -NotePropertyName 'RMMTask' -NotePropertyValue 0.0 -Force
    $RMMTickets = Read-AutoTaskTickets -LastActionFromDate $earliestDate -TitleBeginsWith RMM
    if ($RMMTickets) {
        foreach ($ticket in $RMMTickets) {
            $items = $timeEntries | Where-Object ticketid -eq $ticket.id
            foreach ($item in $items) {
                $item.RMMTicket = $item.hoursWorked
                Write-debug "Build-AutoTaskRMMTickets: found RMM ticket time entry $($item.hoursworked)  on $($RMMTickets.id) $($RMMTickets.title) "
            }   
        }
    }

    foreach ($rmm in $RMMtaskCodes) {
        $RMMtasks = $timeEntries | Where-Object BillingCodeID -eq $rmm
        foreach ($task in $RMMtasks) {
            $task.RMMTask = $task.hoursWorked
        }        
    }

    return $timeEntries
}

function Read-CompanyAlert() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]
        $AlertTypeID = 1,
        [Parameter(Mandatory = $true)]
        [string]
        $CompanyID

    )
    Write-Host "Polling Autotask for CompanyID $CompanyID and AlertTypeID $AlertTypeID  "
    $u = Invoke-AutoTaskAPI -entityName 'v1.0/CompanyAlerts' -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""alertTypeID"",""value"":""$AlertTypeID""},{""op"":""eq"",""Field"":""CompanyID"",""value"":""$CompanyID""}" # -Verbose
    if ($u) {
        $u[0].alertText
    }
}

function Read-CompanyChildAlerts() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [int]
        $CompanyID
    )
    Write-Host "Polling Autotask for CompanyID $CompanyID for all its alerts"
    $u = Invoke-AutoTaskAPI  -UrlFixedSuffix "v1.0/Companies/$CompanyID/Alerts" 
    if ($u) {
        $u.items
    }
 
}

<#
.SYNOPSIS
Update Autotask Companies with specific information

.DESCRIPTION
Update Autotask Companies with Primary/ Secondary assignments
If a field parameter is "", then NO ACTION will be taken
If secondary field parameter is "", or is "null", then any secondary assignment for the company will be removed
the company is updated with Primary and Secondary assignments

company(s) need to be identified either by ID or EXEACT company name, else they will not be updated

.PARAMETER CompanyID
the unique autotask ID

.PARAMETER CompanyName
Alternate to ID, the EXACT name of the company in autotask

.PARAMETER Manager
Parameter description

.PARAMETER Classification
Parameter description

.PARAMETER Branch

.EXAMPLE
Set-AutoTaskCompanies -CompanyID 0  -branch "Tauranga - Kiss I.T"  -Verbose
Set-AutoTaskCompanies -CompanyID 0  -branch "Tauranga"  -Verbose
Set-AutoTaskCompanies -CompanyID 0  -branch 29682914
Set-AutoTaskCompanies -CompanyName "Kiss IT"  -branch 29682914
Set-AutoTaskCompanies -CompanyName "Kiss IT"  -branch 29682914 -Manager "sean Macey"
Set-AutoTaskCompanies -CompanyName "Kiss IT"  -branch 29682914 -Manager 30761525 -Classification 13
Set-AutoTaskCompanies -CompanyName "Kiss IT"  -branch 29682914 -Manager "sean Macey" -Classification 13
Set-AutoTaskCompanies -CompanyName "Kiss IT"  -branch 29682914 -Manager "sean Macey" -Classification Residential
Set-AutoTaskCompanies -CompanyName "Kiss IT"  -branch Tauranga  -Manager "sean Macey" -Classification Residential
$CSV | Set-AutoTaskCompanies
.NOTES
General notes
#>#
function Set-AutoTaskCompanies() {
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
        if ($Classification -and ($Classification -ne "null")) { $classes = Read-AutoTaskCompanyClassificationIcons }
        if ($Manager -and ($Manager -ne "null")) { $Engineers = Read-AutoTaskEngineers -isActive }
        if ($Branch -and ($Branch -ne "null")) { $Branches = ( Invoke-AutoTaskAPIREST -Method GET -url '/V1.0/UserDefinedFieldListItems/query?search={"filter":[{"op":"eq","field":"udfFieldId","value":"29682941"}]}' ).items }
        #  $ipatch = 0
        # $patchObj =@()
    
        <#
      UserDefinedFieldDefinitions  . Branch => id = 29682941 (datatype 3)
      #>
    }
    process {
 

        write-verbose "Set-AutotaskCompanies CompanyID to process = $CompanyID"

        # if ($Manager -eq "null"){ $Manager = ""}
        # if ($CompanyType -eq "null") { $CompanyType = "" }
        if (($CompanyID -eq -1) -and $CompanyName) {
            #   if (!$CompanyID) {
            Write-Verbose "Set-AutotaskCompanies About to check by name of Company $CompanyName : Comnpany ID= $CompanyID"
            $res = read-autotaskCompanies -CompanyName $CompanyName -exactNameMatch -DontExpandChildIDFields
            if ($Res) { $CompanyID = $res.ID }

        }
        
        If ($CompanyID -eq -1) { return }      
        foreach ($anID in $CompanyID) {
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
                        throw "Set-AutoTaskCompanies: Can not fully update CompanyID $anID : could not find Engineer/Manager in autotask matching $Manager "
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
                    #$obj.Classification = ""}
                    $obj | Add-Member -NotePropertyName "classification" -NotePropertyValue ""

                }
                else {
                    $val = ""
                    $res = $classes | Where-Object name -like $Classification
                    if ($res) {
                        $val = $res.id
                        $obj.id = $anID
                        $obj | Add-Member -NotePropertyName "classification" -NotePropertyValue $val
                        else {
                            throw "Set-AutoTaskCompanies: Can not fully update CompanyID $anID : could not find classification in autotask matching $Classification "

                        } 
                         
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
                Invoke-AutoTaskAPIREST -url 'V1.0/Companies' -Method PATCH -Body $json | Out-Null
              
                # $patchObj += $Obj
                # if ($patchObj.count -gt 200){
                #     $json = ($patchObj | ConvertTo-Json -Compress).trim("[").trim("]")
                #     Write-verbose " Set-AutotaskCompany update Json body $json"
                #     Invoke-AutoTaskAPIREST -url 'V1.0/Companies' -Method PATCH -Body $json | Out-Null
                #     $patchObj = @()
                # }

            }

            
        }
    }
    
    end {
        # if ($patchObj.count -gt 0){
        #     $json = ($patchObj | ConvertTo-Json -Compress).trim("[").trim("]")
        #     Write-verbose " Set-AutotaskCompany update Json body $json"
        #     Invoke-AutoTaskAPIREST -url 'V1.0/Companies' -Method PATCH -Body $json | Out-Null
        # }
    }



}

<#
.SYNOPSIS
Update Autotask Companies with Primary/ Secondary assignments

.DESCRIPTION
Update Autotask Companies with Primary/ Secondary assignments
If primary field parameter is "", or is "null", then any primary assignment for the company will be removed
If secondary field parameter is "", or is "null", then any secondary assignment for the company will be removed
the company is updated with Primary and Secondary assignments

company(s) need to be identified either by ID or EXEACT company name, else they will not be updated

.PARAMETER CompanyID
the unique autotask ID

.PARAMETER CompanyName
Alternate to ID, the EXACT name of the company in autotask

.PARAMETER Primary
Parameter description

.PARAMETER Secondary
Parameter description

.EXAMPLE
$eng = import-csv .\PrimaryEngineers.csv
$eng | Set-AutoTaskPrimaryEngineers

or 
Set-AutoTaskPrimaryEngineers -CompanyName "Matamata Medical Center" -Primary "sean" -Secondary "Antony"

or Set-AutoTaskPrimaryEngineers -ID 29762990 -Primary sean -Secondary null

.NOTES
General notes
#>#
function Set-AutoTaskPrimaryEngineers() {
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

        if (($CompanyID -eq -1) -and $CompanyName) {
            $res = read-autotaskCompanies -CompanyName $CompanyName -exactNameMatch -DontExpandChildIDFields
            if ($res) { $CompanyID = $res[0].id }
        }
        If ($CompanyID -eq -1) { return }

        if ($primary -eq "null") { $primary = "" }
        if ($secondary -eq "null") { $Secondary = "" }
        foreach ($anID in $CompanyID) {
            write-host "CompanyID to process = $anID"

            $ChildAlerts = Read-CompanyChildAlerts -CompanyID $anID
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
                        $alert.alertText = "Primary Engineer: $primary`n"
                    }
                    if ($secondary -and ($x -ne 2)) {
                        $alert.alertText = $alert.alertText + "Secondary Engineer: $secondary"
                    }


                    if ($atemp) {
                        Write-Verbose "write-companyPrimary: alerttypeid:$x found extra text $atemp"
                        $alert.alertText = $alert.alertText.trim() + "`n" + "$atemp"
                    }
                    $alert.alertText = $alert.alertText -replace '^(\n)*', "" #-replace "^(`n",""
                    if ($alert.alertText) {
                        #the alert exists - so update it
                        $json = $alert | ConvertTo-Json 
                        write-verbose "write-CompanyPrimary alertTypeID:$x Updating Primary for $anID"
                        Invoke-AutoTaskAPIREST -url ('V1.0/Companies/' + $anID + '/Alerts') -Method PUT -Body $json | Out-Null
                        # $jsontxt += $json
                    }
                    elseif ($alert.id) {
                        #there is no needed alertText, so DELETE the alert
                        write-verbose "write-CompanyPrimary alertTypeID:$x Deleting Primary for $anID"
                        Invoke-AutoTaskAPIREST -url ('V1.0/Companies/' + $anID + '/Alerts/' + $alert.id) -Method DELETE  | Out-Null
                        #  $jsontxt += $json

                    }

                }
                else {
                    if (($Primary -or $Secondary) -and ($x -ne 2)) {
                        #creating a new alert
                        Write-Verbose "write-CompanyPrimary alertTypeID:$x creating a NEW alert record"
                        $alert = [PSCustomObject]@{
                            alertText   = "Primary Engineer: $primary`nSecondary Engineer:$secondary"
                            alertTypeID = $x
                            companyID   = $anID
                        } 
                        $json = $alert | ConvertTo-Json
                        Invoke-AutoTaskAPIREST -url ('V1.0/Companies/' + $anID + '/Alerts') -Method POST -Body $json | Out-Null
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

function Read-AutoTaskPrimaryEngineers() {
    <#
    .SYNOPSIS
    get  primary and secondary technician assignments to Customers
    
    .DESCRIPTION
    will return an object array of
    AccountID  #CustomerID
    Prime   #Primary Tech
    Secondary #Secondary Tech
    
    .EXAMPLE
    Read-PrimeEngineers()
    
    .NOTES
    General notes
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
    $u = Invoke-AutoTaskAPI -entityName 'v1.0/CompanyAlerts' -SearchFirstBy Nothing -SearchFurtherBy "{""op"":""eq"",""Field"":""alertTypeID"",""value"":""$alertTypeID""},{""op"":""contains"",""Field"":""alertText"",""value"":""primary""}" # -Verbose
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
            $classifications = Read-AutoTaskCompanyClassificationIcons
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
               
                $company = Read-AutoTaskCompanies -id $assignedTech.CompanyID -DontExpandChildIDFields
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


function Read-AutoTaskEngineers() {
    <#
    .SYNOPSIS
    provides a list of active Engineers, and assigns items such as 8 hours per day
    - Leo's time is Daily Hrs is Hard Coded as 4 hours here
    
    .DESCRIPTION
        provides a list of active Engineers, and assigns items such as 8 hours per day
    - Leo's time is Daily Hrs is Hard Coded as 4 hours here
    
    .PARAMETER IncludeAllFieds
    if used then all fields from autotask will be returnd
    
    .PARAMETER DeafultDailyHrs
    only relevant if autotask does not have a HR entry for the resource,
    if an HR entry exists then the builddailyhrs will use those figures to calculate each dqaily expectedHrs
    default is 0
    each Engineer will be expected to do these hours of normal work

    .EXAMPLE
    Read-AutoTaskEngineers
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (

        [switch]
        $IncludeAllFieds = $false,
        # [single]
        # $DeafultDailyHrs = 0,
        [switch]
        $isActive = $false
    )
    write-Host "Polling Autotask about Resources (Engineers)"
    $includeFields = $null
    if (!$IncludeAllFieds) {
        $includeFields = "id", "userName", "firstName", "LastName", "resourceType", "isActive", "mobilePhone", "payrollIdentifier", "userType", "title", "hireDate"
        $IncludeFields = ('"{0}"' -f ($includeFields -join '","')).replace('""', '"')
    }



    $result = Invoke-AutoTaskAPI -entityName 'v1.0/Resources' -includeFields $includeFields -SearchFurtherBy '{"op":"noteq","Field":"userType","value":"17"}'  -isActive:$isActive
    $result | Add-Member -NotePropertyName FullName -NotePropertyValue ""
    # $result | Add-Member -NotePropertyName dailyHrs -NotePropertyValue $DeafultDailyHrs
    $DailyAvialabilities = Invoke-AutoTaskAPI -entityName 'v1.0/ResourceDailyAvailabilities'

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
    write-Host "DONE Polling Autotask about Resources (Engineers)" -ForegroundColor Green
    return $result
}


# function Read-AutotaskQueues() {
#     [CmdletBinding()]
#     param (
#         # [Parameter()]
#         # [TypeName]
#         # $ParameterName
#     )

#     $result = Invoke-AutoTaskAPI -entityName 'v1.0/Resources' #-includeFields $includeFields -SearchFurtherBy '{"op":"noteq","Field":"userType","value":"17"}'  -isActive:$isActive
 
# }





function Read-AutoTaskTimeEntries() {
    <#
    .SYNOPSIS
    polls autotask for a list of time entries
    
    .DESCRIPTION
    collects infromation on timesheet entries
    adds summaries of some time of work flows (productive vs internal etc leave) 
    
    .PARAMETER LastxMonths
    how many months back to start the import from
    
    .PARAMETER nonStatCodes
    these BillingCodeID will be assessed as notStat => such as Holiday or Sick Leave
    
    .PARAMETER afterHrsBillingCodes
    these billing codes will be recognised as afterhours
    
    .EXAMPLE
    $i = Read-AutoTaskTimeEntries -LastxMonths 3 
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        # Parameter help The number of months earlier than now, from which to start pulling the time sheeting data from
        [Parameter()]
        [int]
        $LastxMonths = 3,

        #ticket codes
        [int[]]
        $afterHrsBillingCodes = (29683343, 29737351),
        # Parameter help these BillingCodes are such as Sick Leave or Holidays and thus shouldn't be measured during productivity %

        [int[]]
        $nonBillableCodes = (29682861), #Non Billable Support

        #------------internal codes
        #[Parameter(AttributeValues)]
        [int[]]
        $LeaveCodes = (91206, 29718729),
       
        # Parameter help these BillingCodes are such as Sick Leave or Holidays and thus shouldn't be measured during productivity %
        #[Parameter(AttributeValues)]
        [int[]]
        $SickCodes = (91207),
        
        [int[]]
        $teabreakCodes = (91209),

        [int[]]
        $TrainingCodes = (29683344), #, training
  
        [int[]]
        $ProductiveCodes = (29711172, 29712660, 29713657, 29737360, 29718730, 29737360), #Second Level Support, RMM, presales, research, renewals

        [int]
        $RMMCode = 29712660


        
    )

    write-Host "Polling AutoTask for TimeEntries, and formating the results"
    $CURRENTDATE = GET-DATE -Hour 0 -Minute 0 -Second 0
    $Monthstart = $CURRENTDATE.AddMonths(-$LastxMonths)
    #$Monthstart = $CURRENTDATE.AddDays(-7)
    $MonthStartSTr = $Monthstart.ToString("yyyy-MM-ddTHH:mm:ss")

    # $FIRSTDAYOFMONTH = GET-DATE $Monthstart -Day 1
    # $LASTDAYOFMONTH = GET-DATE $FIRSTDAYOFMONTH.AddMonths(1).AddSeconds(-1)

    $includefields = "id", "billingCodeID", "taskID", "ticketID", "timeEntryType", "startDateTime", "endDateTime", "resourceID", "isNonBillable", "hoursWorked", "hoursToBill", "offSetHours", "dateWorked"

    $searchby = '{"op":"gte","Field":"dateWorked","value":"' + $MonthStartSTr + '"}'

    $timeentries = Invoke-AutoTaskAPI -entityName 'v1.0/TimeEntries' -SearchFurtherBy $searchby -SearchFirstBy Nothing -includeFields $includefields 
    Convert-ObjArrayDateTimesToSearchableStrings -obj $timeentries 
   

    write-verbose "Read-AutoTaskTimeEntries count = $($timeentries.count)"
    # Now provide calculate Columns to assist with stats
    $timeentries | Add-Member -NotePropertyName 'OADate' -NotePropertyValue 0.0
    $timeentries | Add-Member -NotePropertyName 'kissWorkType' -NotePropertyValue ""
    #    $timeentries | Add-Member -NotePropertyName 'isAfterHrs' -NotePropertyValue "" 
    # $timeentries | Add-Member -NotePropertyName 'TeaBreaks' -NotePropertyValue ""
 
    $timeentries | Add-Member -NotePropertyName 'HrsTicketBillableNormalHrs' -NotePropertyValue 0.0  
    $timeentries | Add-Member -NotePropertyName 'HrsTicketNonBillableNormalHrs' -NotePropertyValue 0.0 
    $timeentries | Add-Member -NotePropertyName 'HrsTicketBillableAfterHrs' -NotePropertyValue 0.0 
    $timeentries | Add-Member -NotePropertyName 'HrsTicketNonBillableAfterHrs' -NotePropertyValue 0.0 
    $timeentries | Add-Member -NotePropertyName 'Ticket' -NotePropertyValue 0.0 

    $timeentries | Add-Member -NotePropertyName 'HrsLeave' -NotePropertyValue 0.0 
    $timeentries | Add-Member -NotePropertyName 'HrsSick' -NotePropertyValue 0.0
    $timeentries | Add-Member -NotePropertyName 'HrsTeaBreaks' -NotePropertyValue 0.0
    $timeentries | Add-Member -NotePropertyName 'HrsTraining' -NotePropertyValue 0.0
    $timeentries | Add-Member -NotePropertyName 'HrsInternalProd' -NotePropertyValue 0.0
    $timeentries | Add-Member -NotePropertyName 'HrsInternalOther' -NotePropertyValue 0.0 
    $timeentries | Add-Member -NotePropertyName 'AfterHours' -NotePropertyValue 0.0
    
    
  
    #create a numerically sortable date field
    foreach ($i in $timeentries) {
        $i.OADate = ([datetime]$i.dateWorked).ToOADate()
    }


    #---------------
    #Process the Ticket (customer related) time entries
    $subitems = $timeentries | Where-Object ticketID 
    if ($subitems) {
    
        $items = $subitems | Where-Object { (($_.isNonBillable -eq $true) -or ($_.billingCodeID -in $nonBillableCodes)) }
        if ($items) {
            $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Ticket-NonBillable-NormalHrs" -Force
            foreach ($item in $items) {
                $item.HrsTicketNonBillableNormalhrs = $item.hoursWorked
                $item.isNonBillable = $true
                $item.ticket = $item.hoursWorked
            } 
        }
     
        $items = $subitems | Where-Object { ($_.isNonBillable -ne $true) }
        if ($items) {
            $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Ticket-Billable-NormalHrs" -Force
            foreach ($item in $items) {
                $item.HrsTicketBillableNormalHrs = $item.hoursWorked
                $item.ticket = $item.hoursWorked
            }   
        }

        #identify the afterhours billable
        $items = $subitems | Where-Object { ($_.billingCodeID -in $afterHrsBillingCodes) -and ($_.isNonBillable -ne $true) }
        if ($items) {
            $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Ticket-Billable-AfterHrs" -Force
            foreach ($item in $items) {
                $item.HrsTicketBillableAfterHrs = $item.hoursWorked
                $item.HrsTicketBillableNormalHrs = 0
                $item.ticket = $item.hoursWorked
                $item.afterhours = $item.hoursWorked
            }

        }
        #identify the afterhours nonbillable
        $items = $subitems | Where-Object { ($_.billingCodeID -in $afterHrsBillingCodes) -and ($_.isNonBillable -eq $true) }
        if ($items) {
            $items | Add-Member -type NoteProperty   -Name 'kissWorkType' -Value "Ticket-Non-Billable-AfterHrs" -Force
            foreach ($item in $items) {
                $item.HrsTicketNonBillableAfterHrs = $item.hoursWorked
                $item.HrsTicketNonBillableNormalHrs = 0
                $item.ticket = $item.hoursWorked
                $item.afterhours = $item.hoursWorked
            }

        }
    }

    else { Write-Verbose "No ticket items found in timesheet entries" }

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


    Build-AutoTaskInternalTicketsTime $timeentries | Out-Null
    Build-AutoTaskRMMTime $timeentries | Out-Null
    write-Host "DONE polling AutoTask for TimeEntries, and formating the results" -foregroundcolor green

    return $timeentries
   
}


function export-KissAtCompanies() {
    <#
    .SYNOPSIS
    create a CSV file containing a list of companies
    
    .DESCRIPTION
    Long description
    
   
    .PARAMETER exportType
    CSV or JSON
    Deafult == CSV
    JSON does not work yet
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
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
    write-host "Export-KissAtCompanies will take about 3 minutes to run!"
    switch ($exportType) {
        "CSV" {
            write-host "export-KissAtCompanies =>Exporting ClassificationIcons"
            Invoke-AutoTaskAPI -entityName 'v1.0/ClassificationIcons' -includeFields "id", "name" | export-csv "$($path)KissAtClassificationIcons.csv" -NoTypeInformation -Force
            write-host "export-KissAtCompanies =>Exporting Companies"
            Read-AutoTaskCompanies -includeInactive -GetEngineers | export-csv "$($path)KissAtCompanies.csv" -NoTypeInformation -Force
        }
        default {

        }

    }
    write-host "Done Export-KissAtCompanies" -ForegroundColor green
}



function IsWorkingDay() {
    <#
    .SYNOPSIS
    determines weather a date is in the wokring day or weekend
    true == working day
    false == weekend
    
    .DESCRIPTION
    Long description
    
    .PARAMETER date
    get-date
    
    .EXAMPLE
    (IsWorkingDay($result.workDate)
    
    .NOTES
    General notes
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


function Build-AutotaskDailyTimeStats {
    <#
    .SYNOPSIS
    calculate daily summary for each technician that is time sheeting
    requires the timeEntries object array to t=be parsed to it
    - this does not use inline processing, the timeentries must be passed as a paramneter object array
    
    .DESCRIPTION
    Long description
    creates daily expected hours, which is the greater of (normal ours worked less Leave and Sick) Or each Tech's ecpected daily hours
    
    .PARAMETER TimeEntries
    AN array of tiome entries (generated by Read-AutoTaskTimeEntries )
    
    .EXAMPLE
    Build-AutotaskDailyTimeStats -TimeEntries $timeEntries
    
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Position = 0, Mandatory = $true)]   
        [PSCustomObject]        $TimeEntries,
        [datetime]$UntilDate = (get-date) # check u timesheeted days for resources from earliest in toimesheet until this time - so ignore leave requests and future bookings when filling gaps
    )



    $allresources = Read-AutoTaskEngineers
    $Resources = $allresources | Where-Object { ($_.id -in $TimeEntries.resourceID) }  ## gets resources in time entries
    # $ResourcesThatcouldhavetimesheeted = $allresources | Where-Object {($_.DailyAvailabilities.MondayAvailableHours -or $_.DailyAvailabilities.TuesdayAvailableHours -or $_.DailyAvailabilities.WednesdayAvailableHours -or $_.DailyAvailabilities.ThursdayAvailableHours -or $_.DailyAvailabilities.FridayAvailableHours -or $_.DailyAvailabilities.SaturdayAvailableHours -or $_.DailyAvailabilities.SundayAvailableHours  )}
    $Resources += $allresources | Where-Object { ($_.isActive) -and ($_.DailyAvailabilities.MondayAvailableHours -or $_.DailyAvailabilities.TuesdayAvailableHours -or $_.DailyAvailabilities.WednesdayAvailableHours -or $_.DailyAvailabilities.ThursdayAvailableHours -or $_.DailyAvailabilities.FridayAvailableHours -or $_.DailyAvailabilities.SaturdayAvailableHours -or $_.DailyAvailabilities.SundayAvailableHours  ) }
    #$resourcesThatShouldTimeSheet = $Resources | Select-Object * -Unique
    
    write-verbose "Build-AutotaskDailyTimeStats: Resources that are expected to be timesheeting $($resources.username -join (', '))"
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

        [psobject[]]$OneResourceDates = $null
        #Find all resources which have time entries
        $Resource = $Resources | Where-Object { ($_.id -eq $gp.name) } | Select-Object -First 1
        

        $techDays = $gp.Group | Group-Object dateworked
        foreach ($techDay in $techDays) {
            $result = [PSCustomObject]@{
                resourceID                         = $Resource.id
                Resource                           = $Resource.username
                workDate                           = [string]$techday.name
                HoursExpectedPerDay                = $Resource.dailyHrs  
                hoursWorked                        = ($techDay.group | Measure-Object -Property hoursWorked -sum).sum
                hrsTicketBIllableNormalHrs         = ($techDay.group | Measure-Object -Property hrsTicketBIllableNormalHrs -sum).sum
                hrsTicketBIllableAfterHrs          = ($techDay.group | Measure-Object -Property hrsTicketBIllableAfterHrs -sum).sum
                hrsTicketNonBIllableNormalHrs      = ($techDay.group | Measure-Object -Property hrsTicketNonBIllableNormalHrs -sum).sum
                hrsTicketNonBIllableAfterHrs       = ($techDay.group | Measure-Object -Property hrsTicketNonBIllableAfterHrs -sum).sum
                HrsLeave                           = ($techDay.group | Measure-Object -Property HrsLeave -sum).sum
                HrsSick                            = ($techDay.group | Measure-Object -Property HrsSick -sum).sum
                HrsTeaBreaks                       = ($techDay.group | Measure-Object -Property HrsTeaBreaks -sum).sum
                HrsTraining                        = ($techDay.group | Measure-Object -Property HrsTraining -sum).sum
                HrsInternalProd                    = ($techDay.group | Measure-Object -Property HrsInternalProd -sum).sum
                HrsInternalOther                   = ($techDay.group | Measure-Object -Property HrsInternalOther -sum).sum
                InternalTicketBillableNormalHrs    = ($techDay.group | Measure-Object -Property InternalTicketBillableNormalHrs -sum).sum
                InternalTicketBillableAftHrs       = ($techDay.group | Measure-Object -Property InternalTicketBillableAftHrs -sum).sum
                InternalTicketNonBillableNormalHrs = ($techDay.group | Measure-Object -Property InternalTicketNonBillableNormalHrs -sum).sum
                InternalTicketNonBillableAftHrs    = ($techDay.group | Measure-Object -Property InternalTicketNonBillableAftHrs -sum).sum
                InternalTicketTotal                = ($techDay.group | Measure-Object -Property InternalTicket -sum).sum
                TicketTotal                        = ($techDay.group | Measure-Object -Property Ticket -sum).sum
                AfterHours                         = ($techDay.group | Measure-Object -Property AfterHours -sum).sum
                RMMTicket                          = ($techDay.group | Measure-Object -Property RMMTicket -sum).sum
                RMMTask                            = ($techDay.group | Measure-Object -Property RMMTask -sum).sum
            }
    
            if ($Resource.DailyAvailabilities) {
                $DayNum = ([datetime]($Result.workDate)).DayOfWeek.value__
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
                
                Write-Debug "BUild-AutoTaskDailyTimeStats: Day hours for $($result.Resource) on day $daynum is $($result.HoursExpectedPerDay) "
            }
            else {
                Write-Debug "BUild-AutoTaskDailyTimeStats: Day hours for $($result.Resource) were not found"
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

function export-KissAtTimerecords() {
    <#
    .SYNOPSIS
    create mulotipe CSV or JSON files, that can be used with powerBI etc
    will create these in the directory the the script is run fdrom
    
    .DESCRIPTION
    Long description
    
    .PARAMETER LastxMonths
    default is 3 months
    how many monbths back to retreive the data from
    
    .PARAMETER exportType
    the default is CSV
    either CSV or JSON (S+CSV work, JSON needs review)
    
    .EXAMPLE
    w:
    w:\autotask\
    export-KissAtTimerecords() -LastxMonths CSV
    
    .NOTES
    General notes
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
    Write-Host "export-KissAtTimerecords: will take some time to run"
    write-host " export-KissAtCompanies =>preparing Time Entries"

    $i = Read-AutoTaskTimeEntries -LastxMonths $LastxMonths

    $earliestDate = ($i | Measure-Object dateWorked -min).Minimum
    write-host " export-KissAtCompanies =>preparing Ticket Details"

    $Tickets = Read-AutoTaskTickets -LastActionFromDate $earliestDate
    if ($path) { $path = "$path\\" }

    switch ($exportType) {
        "CSV" {
            write-host " export-KissAtCompanies =>Billing Codes"

            Invoke-AutoTaskAPI -entityName 'v1.0/BillingCodes' | Export-csv "$($path)KissBillingCodes.csv" -NoTypeInformation -Force
            write-host " export-KissAtCompanies =>Resources (Engineers) and timeEntries"
            Read-AutoTaskEngineers | export-csv "$($path)KissEngineers.csv" -NoTypeInformation -Force
            $i | export-csv "$($path)KissTimeEntries.csv" -NoTypeInformation -Force
            write-host " export-KissAtCompanies =>DailyTime Stats and tickets"
            Build-AutotaskDailyTimeStats -TimeEntries $i | Export-Csv "$($path)KissDaily.csv" -NoTypeInformation -Force
            $Tickets | Export-Csv "$($path)KissTickets.csv" -NoTypeInformation -Force

            #  Invoke-AutoTaskAPI -entityName 'v1.0/ResourceTimeOffBalances' | Export-csv ResourceTimeOffBalances.csv -NoTypeInformation

            #Holiday and Holidayset records not in use
        }
        default {
            Invoke-AutoTaskAPI -entityName 'v1.0/BillingCodes' | ConvertTo-Json | Out-File -FilePath KissBillingCodes.json -Force
            Read-AutoTaskEngineers | ConvertTo-Json | Out-File -FilePath KissEngineers.json -Force
            # $i = Read-AutoTaskTimeEntries -LastxMonths $LastxMonths
            $i | ConvertTo-Json | Out-File -FilePath  KissTimeEntries.json -Force
            Build-AutotaskDailyTimeStats -TimeEntries $i | ConvertTo-Json | Out-File -FilePath  KissEnginerDailies.json -Force
            $Tickets | ConvertTo-Json | Out-File -FilePath  KissTickets.json -Force
            #    Invoke-AutoTaskAPI -entityName 'v1.0/ResourceTimeOffBalances' | Out-File -FilePath ResourceTimeOffBalances.json

        }

    }
    write-host "Done export-KissAtCompanies" -ForegroundColor green
}

function export-KissATTickets() {
    param (
        [Parameter(Mandatory = $false)]
        [int]
        $WhereLastActionOccurWithinLastMonths = 0,
        [string]$path
    )

    if ($path) { $path = "$path\\" }
    New-Item -ItemType Directory -Name data -ErrorAction SilentlyContinue | Out-Null
    if ($LastActionAfter -gt 0) {
        Read-AutoTaskTickets -LastxMonths $WhereLastActionOccurWithinLastMonths | Export-csv "$($path)TicketsActioned.csv" -NoTypeInformation -Force
    }
    Read-AutoTaskTickets -IncludeAllNonComplete | Export-csv "$($path)TicketsNotCompleted.csv" -NoTypeInformation -Force

}


function Set-loginAutotask() {
    <#
    .SYNOPSIS
    āllows automatic connection to the AutoTask API
    
    .DESCRIPTION
       checks credentials and APIID, then saves them encrypted within a file opn the users home\kiss-atapi path
    this function does accept inline, but this is not needed.
    the best practice is to enter the values when prompted and they will be as SECURE strings (no one can see them..)
    
    .PARAMETER l_username
    API user name (usually an email address and NOt a firstname.lastname format). this is a GLOBALLY useable username
    
    .PARAMETER l_pass
    password for the API user
    saved as an encrypted file, and used a secure string
    
    .PARAMETER l_apiid
    API ID for the API user
    saved as an encrypted file, and used a secure string
    
    .EXAMPLE
     Set-loginAutotask
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
        atapi    = ''#ConvertFrom-SecureString -SecureString $l_Apiid
        UserName = ''#"$apiusername"
        Secret   = '' #ConvertFrom-SecureString -SecureString $l_secret
        url      = ''# "$($r.url)"
    }

    if (!(Test-Path -Path $kissATAPIpath)) {
        new-item -Path $home -Name kiss-atapi -ItemType Directory
        Write-Host "Created a new Directory called $($home)\kiss-atapi" 
    }
    else {
        if (test-path -path "$kissATAPIpath\$kissATAPIfile" ) {
            $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
            if ($jsn) {
                write-host "there was a prexisting saved login of $jsn"
                $r = $jsn  | ConvertFrom-Json 
            }
            if ($r.url -and $r.secret -and $r.username -and $r.atapi) {
                #saved data exists and is valid , so import it
                $saveobj = $r
            }
        }
    }
    if ($l_username) { $saveobj.UserName = $l_username }
    if ($l_pass) {
        $saveobj.Secret = $l_pass  | ConvertTo-SecureString -AsPlainText -Force # this converts it to secure string
        $saveobj.Secret = $saveobj.Secret | ConvertFrom-SecureString  # this encrypts it
    }

    if ($l_apiid) { $saveobj.atapi = $l_apiid }

    write-verbose "user = $($saveobj.UserName)"
    write-verbose "pass = hiiden"#$($saveobj.Secret)"
    write-verbose "atapi = $($saveobj.atapi)"
    

    if ($saveobj.userName) {
        Write-Host "there is already definition saved : for $($saveobj.UserName)"
        write-host "  If you wish to keep the old settings, then just hit return on that field without entering anything"
    }

 
  
    $i = read-host -Prompt "Enter a new API USER "
    
    write-host "now checking with the remote autotask API...."
    if ($i) { $saveobj.username = $i }
    $r = Invoke-RestMethod -Uri "http://webservices.autotask.net/atservicesrest/v1.0/zoneInformation?user=$($saveobj.username)"
    
    if ($r.url) {
        write-host "$l_username will use the following autotask API intergface:   $($r.url)"
        $saveobj.url = $r.url
    }
    else {
        write-host "$l_username is not a valid user within the autotask API (or the autotask API could not be contacted at this time), please retry"
        return
    }

    $i = read-host -Prompt "Enter the USER's password (Alphanumerical and special)" -AsSecureString -ErrorAction SilentlyContinue
    if ($i.length -gt 0) {
        $saveobj.Secret = $i | ConvertFrom-SecureString  # makes it a secure string, then encrypts it
    }
   

    $i = read-host -Prompt "Enter the AT-API-ID  {alphanumerical}" #-AsSecureString -ErrorAction SilentlyContinue
    if ($i.length -gt 0) { $saveobj.atapi = $i }


 

    $jsn2 = ConvertTo-Json $saveobj

    Write-Verbose "Json = $jsn2"
    Write-verbose "setting json content to $kissATAPIpath\$kissATAPIfile"

    # switch ($force) {
        
    #     $true { Set-Content "$kissATAPIpath\$kissATAPIfile" -Value $jsn2; break }
    #     ((Test-AutoTaskConnection -LoginInfo $saveobj) ) { Set-Content "$kissATAPIpath\$kissATAPIfile" -Value $jsn2; break }

    #     default {}
    # }

    $testresult = Test-AutoTaskConnection -LoginInfo $saveobj
    $testresult
    if ($testresult) {
        write-Host "Set-LoginAutotask:Saving the entered Login becasue the test connection was successfull"
    }
    elseif ($force -eq $true) {
        write-Host "Set-LoginAutotask: even though this login did not work, we are saving it because the FORCE paramater was used"
    }
    else {
        write-Host "Set-LoginAutotask: Not saving a thing, becasue the login didn't work"
        return
    }
    write-host "setcontent on  $kissATAPIpath\$kissATAPIfile   value $jsn2"
    Set-Content "$kissATAPIpath\$kissATAPIfile" -Value $jsn2

}

function Test-AutoTaskConnection {
    [CmdletBinding()]
    param(
        [PSCustomObject]$LoginInfo,
        [switch]$LoginInfoPasswordAsPlainText


 
    )
    if ($LoginInfo) {
        Write-verbose "Test-AutoTaskConnection: will use the parametised LoginInfo  to test, and not the presaved"
            
        
        write-verbose "Test-AutoTaskConnection: Parameters are $($loginfo |ConvertTo-Json)"
        $r = $LoginInfo
        if ($LoginInfoPasswordAsPlainText -ne $true) {
            $Plainpassword = $LoginInfo.Secret | Convertto-SecureString 
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Plainpassword) 
            $r.Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)   
            #  write-verbose "the login that will be tested is  = $($r |ConvertTo-Json)"    
        }
    }

    elseif (test-path -path "$kissATAPIpath\$kissATAPIfile" ) {
        $jsn = Get-Content "$kissATAPIpath\$kissATAPIfile"
        write-verbose "Test-AutoTaskConnection: saved JSON = $jsn"
        if ($jsn) { $r = $jsn | ConvertFrom-Json }
        if ($r.url -and $r.secret -and $r.username -and $r.atapi) {
            #saved data exists and is valid , so import it
            write-host "Test-AutoTaskConnection: will test the connection using credentials for $($r.username)"
            $Plainpassword = $r.Secret | Convertto-SecureString         
            #  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Plainpassword)
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Plainpassword) 
            $r.Secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)       
        }
    }
    else {
        write-host " **** Test-AutoTaskConnection: there were no saved credentials and no parameters supplied"
        Write-Warning "Test-AutoTaskConnection: You must first Set-LoginAutotask and save your APID and credentials"
        return $null
    }

    if (!($r.url) -or !($r.secret) -or !($r.username) -or !($r.atapi)) {
        write-Host "Test-AutoTaskConnection:  At least one of the Authentication parameters are missing"
        return $null
    }

    if (!($r.Secret)) {
        #the password is EMPTY - so return NULL
        Write-host "Test-AutoTaskConnection: Your Login Password was BLANK / NULL so this login will not work" -ForegroundColor Yellow
        return $null
    } 


    write-verbose "the login that will be tested is  = $($r |ConvertTo-Json)"


    try {

        Invoke-AutoTaskAPI -url 'https://webservices6.autotask.net/ATservicesRest/v1.0/Version'  -returnRaw -LoginINfo $r
        Write-host "Test-AutoTaskConnection: Connection to the AutoaTask API was successfull: Your credentials work!" -BackgroundColor Green
        return  $true
 


    }

    catch {
        write-host "Test-AutoTaskConnection: sorry but those credentials did not work"
        write-host "$(($error[0]).Exception.Message)"-ForegroundColor Yellow
        write-host "please try again if you want to change your credentials" -foregroundcolor yellow
        
        return $null
    }
    
}

function Read-AutoTaskTickets {
    [CmdletBinding()]
    param (
        [Parameter()]
        [int[]]
        $CompanyIDs, # =    (29762985 , 0, 1, 29740186 , 29761818, 29762138), #      Imatec Solutions (As Customer), then several Kiss companies
        [DateTime]
        $LastActionFromDate = (Get-date).AddDays(-60), # [dateTime]"2023-01-01T00:00:00",
        [string]
        $TitleContains,
        [string]
        $TitleBeginsWith,
        [string[]]
        $includeFields = $null,
        [switch]$ReturnAllFields = $false,
        [switch]$IncludeAllNonComplete = $null,
        [switch]$DontincludeNonComplete = $false,
        [switch]$DontexpandticketInformation = $false,
        [switch]$whereResourceAssigned,
        [int]$InLastDays = $null,
        [int]
        $loopCount = 40,
        [string]
        $DoSearchBy = $null
        #$LastActionFromDate = (get-date).AddMonths(-3)
    )
    write-host "Read-AutoTaskTickets: polling autotask for ticket information"
    if (!($DontexpandticketInformation)) {
        $ticketinfo = Read-AutotaskTicketInformation
    }
    [int]$u = 0
    [string]$i = $null
    $LastActionFromDateStr = $LastActionFromDate.ToString("yyyy-MM-ddTHH:mm:ss")
    if ($InLastDays) {
        $LastActionFromDateStr = $($(Get-Date).AddDays(0 - $InLastDays)).ToString("yyyy-MM-ddTHH:mm:ss")
    }
    if ($companyIDs.count -gt 0) {
        [string]$cc = $CompanyIDs -join ','
        Write-verbose "Read-AutoTaskTickets companyID searched for are $cc"
        $i = '{"op":"in","Field":"CompanyID","value":[' + $cc + ']}'
        $u = $u + 1
    }
    if ($TitleContains -eq $true) {
        $i = ($i + ',{"op":"contains","Field":"title","value":""' + $TitleContains + '""}').Trim(',')
        $u = $u + 1
    }
    if ($TitleBeginsWith -eq $true) {
        $i = ($i + ',{"op":"beginsWith","Field":"title","value":""' + $TitleBeginsWith + '""}').Trim(',')
        $u = $u + 1
    }
    if ($whereResourceAssigned -eq $true) {
        #  $i = ($i + ',{"op":"exist","Field":"assignedResourceID","value":"null"}').Trim(',')
        $i = ($i + ',{"op":"Exist","Field":"assignedResourceID"}').Trim(',')
        $u = $u + 1
    }
    if ($DontincludeNonComplete -eq $true) {
        $i = ($i + ',{"op":"Exist","Field":"completedDate"}').Trim(',')
        $u = $u + 1
    }

    if ($DoSearchBy) {
        $searchby = $DoSearchBy
    }
    else {
        if ($u -gt 0) {
            #          $searchby = '{"op":"and","items":[{"op":"gte","Field":"lastActivityDate","value":"' + $LastActionFromDateStr + '"}'  + ',' + $i + ']}'
            $searchby = '{"op":"and","items":[{"op":"gte","Field":"lastActivityDate","value":"' + $LastActionFromDateStr + '"}' + ',' + $i + ']}'
        }
        else {
            $searchby = '{"op":"gte","Field":"lastActivityDate","value":"' + $LastActionFromDateStr + '"}' #+ ',' + $i 
        }
    }


    
    # $searchby =$searchby -replace ' ',''


    if ($ReturnAllFields) { $includeFields = $null }
    elseif (!$includeFields) {
        $includeFields = ('id', 'TicketNumber', 'CompanyID', 'completedDate', 'createDate', 'firstResponseDateTime', 'lastActivityDate', 'status', 'tickettype', 'completedDate', 'title', 'assignedResourceID', 'queueid')
    }


    if ($IncludeAllNonComplete -eq $true) {
        # OR two operands so that we can get noncomplete tickets as well as any other Searcth
        #    $searchby = '{"op":"or","items":[{"op":"notExist","Field":"completedDate"}' + ',' + $searchby + ']}'
        $searchby = '{"op":"or","items":[{"op":"and","items":[{"op":"notExist","Field":"completedDate"},{"op":"Exist","Field":"assignedResourceID"}]}' + ',' + $searchby + ']}'
    }

    #write-host $i
    write-verbose "Read-AutoTaskTickets: search by : $searchby"
    $items = Invoke-AutoTaskAPI -entityName 'v1.0/Tickets' -includeFields $includeFields  -SearchFurtherBy $searchby -SearchFirstBy Nothing -LoopCount $loopCount -CheckDuplicatesOf "TicketNumber"
    
    #return $items
    if ($items) {
        if ($ticketinfo) {
            $items | Add-Member -NotePropertyName QueueName -NotePropertyValue "" -Force
            $items | Add-Member -NotePropertyName StatusName -NotePropertyValue "" -Force
            $items | Add-Member -NotePropertyName ResourceName -NotePropertyValue "" -Force
            # $items |Add-Member -NotePropertyName Company -NotePropertyValue "" -Force
            
            $Resources = Read-AutoTaskEngineers
            foreach ($titem in $items) {
                $titem.QueueName = (($ticketinfo.queueID) | Where-Object value -eq $titem.queueID | Select-Object label -first 1).label
                $titem.StatusName = (($ticketinfo.status) | Where-Object value -eq $titem.status | Select-Object label -first 1).label
                if ($titem.assignedResourceID) { $titem.ResourceName = ($Resources  | Where-Object id -eq $titem.assignedResourceID | Select-Object  -first 1).userName }
           
            } 
          


         

        }

        Convert-ObjArrayDateTimesToSearchableStrings $items 
    }
    $items
    write-host "DONE -Read-AutoTaskTickets: polling autotask for ticket information" -ForegroundColor Green
}

function Find-CompaniesInTickets() {
    <#
    .SYNOPSIS
    gets a collection of companies for which the tickets belong
    -- this can take a long time -( eg 3 minutes just for 100 outstanding tickets)
    
    .DESCRIPTION
    Long description
    FYI: in order for some filters to work you SHOULD include id in the search - else the query may never finish
    
    .PARAMETER tickets
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [object[]]
        $tickets
    )
    $companies = $tickets | Group-Object CompanyID
    foreach ($companyID in $companies) {
        $company = (Read-AutoTaskCompanies -id $companyID.Name | Select-Object -First 1)
        $company
        #$tcompanies.Group.Company = "KK"#$companyName
        #  $CompanyID.Group | Add-Member -NotePropertyName Company -NotePropertyValue "$($company.CompanyName)" -Force
    }

}


function Read-AutotaskTicketInformation {
    <#
    .SYNOPSIS
    provides an object listing (SOME) of the known status types usedby tickets
   
    
    .DESCRIPTION


    
    .PARAMETER ExportCSV
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [switch]$ExportCSV
    )

    Write-Host "Read-TicketInformation Polling Autotask for TicketInformation queues, status etc. values "
    $fields = (invoke-AutoTaskAPI -UrlFixedSuffix v1.0//Tickets/entityInformation/fields).fields #(name,picklistvalues[value,label,isactive)

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
