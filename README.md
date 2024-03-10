# AutoTaskRest
A collection of PowerShell scripts, that allow time sheeting and ticket information to be extracted from Autotask

The purpose is to allow extact for analysis and reports,
  * showing the non completed tickets, with an emphasis on normal tickets older than 7 daya , and onHold/Waiting tickets older than 30
  * showing productivity of Staff, and comparing that againts their expected working hours
  * showing where effort was spent on billing and non billable productive work

before this code will connect to autotask and extract data, you must <code>Set-LoginKissAT</code> and enter username, password and APID.
<code>
. ./AutotaskRest.ps1 #to load the scripts inot powershell memory, or you can rename to AutotaskRest.PSM1 and deploy it as a mondule
Set-LoginKissAT
</code>

these details are then encrypted and saved in a user file. the credential and APIID are only decryptable by the user that created them and on the machine they were created on. the password and APIID do not appear as clear text within the module (they are handled as secureStrings)


Most of the functions have inbuilt help- - just the Man command (example below)
<code>
man Get-KissATCompanies
  NAME
  Get-KissATCompanies
  SYNOPSIS
  returns a list of companies (or just one of)
</code>
  
I gleaned information from https://autotask.net/help/DeveloperHelp/Content/APIs/REST/REST_API_Home.htm in order to build these scripts

<h2>How datetime fields are handled</h2>
the API needs to be date local invariant, so the searchable date text date format is used 
EXAMPLE  When making a ContractServiceAdjustments call, the effectiveDate is submitted as <b>2023-10-09T02:00:00.00</b>, that is, 2 AM on October 9. Because the API intakes calls in UTC, if that call is made to a US database (UTC + 5), it would seem to change the effective date to October 8th at 9 PM, due to the time zone conversion.
However, because there is no time field in the UI for service adjustments, we don't convert timezone datetime values for date-only fields, we just set the time portion to midnight and accept the date value.

In the example above, the datetime would be saved in the database as <b>2023-10-08T00:00:00.00</b>.
powershell can create this format  example: <code>$Monthstart.ToString("yyyy-MM-ddTHH:mm:ss")</code>


<h2>Filter operators</h2>
Most calls to the API will require the use of one or more filter operators to indicate the type of query you'd like the API to perform. The table below lists the available operators and their definitions.
You can include user-defined fields (UDFs) in your query. By specifying a udf value of true, you indicate to the API that the field you provide in your query is user-defined. The udf expression must always follow the field expression in the API call. It is not necessary to include the udf value if you are not calling a user-defined field.
<code>
  
  "filter": [
        {
            "op": "SelectedOperator",
            "field": "NameofField",
            "udf": true,
            "value": "DesiredValue"
        }
 </code>


