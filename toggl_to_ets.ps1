#
# Call the script like below
# . .\toggl_to_ets.ps1 -since 2016-11-01 -until 2016-11-15
#

param (
    [string]$Start = [DateTime]::Today.AddDays(-1).ToString('yyyy-MM-dd'),
    [string]$End = $Start,
    [DateTime[]]$IrregularDays = @()
)

#Load configuration from file
. $PSScriptRoot\config.ps1

# Function accepts uername and password, returns a hash to be used as a
#  -Headers parameter for an Invoke-RestMethod or Invoke-WebRequest
function New-BasicAuthHeader {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)] $Username,
        [Parameter(Mandatory = $true)] $Password
    )
    # Doing this step-by-step for illustration; no reason not to reduce it to fewer steps
    # In fact, you can uncomment the following line and remove all the following lines in the function
    # @{ Authorization = "Basic {0}" -f [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $Username,$Password))) }

    # Make "username:password" string
    $UserNameColonPassword = "{0}:{1}" -f $Username, $Password
    # Could also be accomplished like:
    # $UserNameColonPassword = "$($Username):$($Password)"

    # Ensure it's ASCII-encoded
    $InAscii = [Text.Encoding]::ASCII.GetBytes($UserNameColonPassword)

    # Now Base64-encode:
    $InBase64 = [Convert]::ToBase64String($InAscii)

    # The value of the Authorization header is "Basic " and then the Base64-encoded username:password
    $Authorization = "Basic {0}" -f $InBase64
    # Could also be done as:
    # $Authorization = "Basic $InBase64"

    #This hash will be returned as the value of the function and is the Powershell version of the basic auth header
    $BasicAuthHeader = @{ Authorization = $Authorization }

    # Return the header
    $BasicAuthHeader
}

#
# Join entries with the same description and start date
#
function Summarize-Entries {
    param(
        $time_entries
    )

    #Convert regular array to Collection as it should work faster
    $entries = New-Object System.Collections.ArrayList
    $entries.AddRange($time_entries)

    $curr_date = $entries[0].start

    $sum_entries = New-Object System.Collections.ArrayList

    foreach ($entry in $time_entries) {

        if ($curr_date.Date -lt $entry.start.Date) {
            $curr_date = $entry.start
        }
        elseif ($curr_date.Date -ne $entry.start.Date) {
            Write-Host "Exiting with code 365: unexpected date order"
            exit 365
        }

        $same_task = $sum_entries.Where( { $_.description -eq $entry.description -and $_.start.Date -eq $entry.start.Date }, 'First')

        if ($same_task) {
            $same_task[0].dur += $entry.dur
        }
        else {
            $sum_entries.Add($entry)>$null
        }
    }

    $sum_entries
}

#
# Convert given object properties from pure String to System.DateTime
#
function string_to_datetime {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][System.Object] $obj,
        [Parameter(Mandatory = $true)][System.String[]] $prop_names
    )

    foreach ($prop in $prop_names) {
        $obj.$prop = [datetime]::ParseExact($obj.$prop, 'yyyy-MM-ddTHH:mm:sszzz', [System.Globalization.CultureInfo]::CurrentCulture)
    }

    #Return object
    $obj
}

function PrepareTo-ETS {
    param(
        [Parameter(Mandatory = $true)]
        [System.Object[]]
        $entries
    )

    #Join entries with the same description
    $sum_entries = Summarize-Entries -time_entries $entries

    $template_tasks = Read-EtsExcelTemplate

    foreach ($entry in $sum_entries) {

        #Convert duration from milliseconds to hours
        $entry.dur = [Math]::Round($entry.dur / 3600000, $config.effort_fractional_digits)
        if ($entry.dur -eq 0) {
            Write-Host "Entry was removed as it's too short:" $entry.start ":" $entry.description -ForegroundColor yellow
        }

        $entry.tags = $entry.tags[0]

        #Set 'none' tag if nothing was set
        if ($null -eq $entry.tags) {
            Write-Host "Entry hasn't tags: " $entry.start ":" $entry.description -ForegroundColor red
            $entry.tags = 'none'
        }

        #Prepend tag with "Irregular hours" if needed
        if (
            -not ($config.projects_without_irregular_time -contains $entry.project) -and
            -not ($entry.tags -like "$($config.irregular_time_prefix)*") -and
            (
                $entry.start.Hour -ge $config.irregular_time_start -or
                $entry.start.Hour -lt $config.irregular_time_end -or
                $entry.start.DayofWeek -in @(0,6) -or
                $entry.start.Date -in $IrregularDays
            )
        ) {
            $entry.tags = $config.irregular_time_prefix + $entry.tags
        }

        # If the entry is related to a project, which name in Toggl differs from the one in ETS,
        # change the project name accordingly.
        if ( $entry.project -in $config.project_prefix.Keys ) {
            $entry.project = $config.project_prefix[$entry.project]
        }

        # Workaround for the ETS system bug: no way to delete trailing spaces.
        $project_task = $null
        foreach ($task in $template_tasks)
        {
            if (($task -replace " ", "") -eq ("$($entry.project).$($entry.tags)" -replace " ", ""))
            {
                $project_task = $task
                break
            }
        }
        if ([string]::IsNullOrEmpty($project_task)) {
            Write-Host "Entry has non-existent Project-Task assigned: " $entry.start ":" $entry.description  "|" $entry.project":"$entry.tags -ForegroundColor red
        }

        Add-Member -InputObject $entry -NotePropertyName 'project_task' -NotePropertyValue $project_task
    }

    #Return only entries with non zero duration
    $sum_entries | Where-Object { $_.dur -gt 0 }
}

function Read-EtsExcelTemplate
{
    $excel = New-Object -Com Excel.Application
    $wb = $excel.Workbooks.Open($config.excel_template_path)

    $sh = $wb.Sheets | Where-Object -Property "Name" -EQ "Projects"

    $Tasks = @()
    $i = 2          # Start from the second line, as the first one is a header

    do
    {
        $Tasks += $sh.Cells.Item($i, 1).Value2
        $i++
    } until ($null -eq $sh.Cells.Item($i, 1).Value2)

    $excel.Workbooks.Close()

    return $Tasks
}

#####################################################################
#                           Main code
#####################################################################

$Headers = New-BasicAuthHeader -Username $config.api_token -Password "api_token"

#Authentication and session creation
Invoke-RestMethod "$($config.toggl_api_uri)/sessions" -Method Post -Headers $Headers -SessionVariable toggl_api_session -Verbose > $null

#Get the client data by client name
$client = Invoke-RestMethod "$($config.toggl_api_uri)/me?with_related_data=true" -Method Get -ContentType "application/json" -WebSession $toggl_api_session -Verbose |
    Select-Object -ExpandProperty data |
        Select-Object -ExpandProperty clients |
            Where-Object { $_.name -eq $config.client_name }

#Get projects related to the given client
$uri = "$($config.toggl_api_uri)/clients/" + $client.id + "/projects"

$projects = Invoke-RestMethod $uri -Method Get -ContentType "application/json" -WebSession $toggl_api_session -Verbose

#Convert array of objects to comma separated string of IDs
$project_ids = '';
foreach ($project in $projects) {
    if ($project_ids -ne '') { $project_ids += ',' }
    $project_ids += $project.id
}

#Load time entries
#Results might be splitted to several pages. So we need a loop to load them all.

$page = 1
$loaded = 0
$time_entries = $null

do {
    $uri = "$($config.toggl_reports_api_uri)/details?workspace_id=" + $client.wid + "&project_ids=" + $project_ids + "&since=" + $Start + "&until=" + $End + "&user_agent=api_test&order_field=date&order_desc=off&page=" + $page
    $result = Invoke-RestMethod $uri -Method Get -ContentType "application/json" -WebSession $toggl_api_session -Verbose
    $time_entries += $result | Select-Object -ExpandProperty data | ForEach-Object -process { string_to_datetime -obj $_ -prop_names @("start", "end") }
    $loaded += $result.per_page
    $page++

} while ($loaded -lt $result.total_count)

$SelectObjectArgs = @{
    'Property' = @(
        @{Name = "Project-Task"; Expression = { $_.project_task } },
        @{Name = "Effort"; Expression = { $_.dur } },
        @{Name = "Description"; Expression = { $_.description } },
        @{Name = "Started Date"; Expression = { $_.start.ToString('d') } },
        @{Name = "Completion Date"; Expression = { $_.end.ToString('d') } }
    )
}

PrepareTo-ETS -entries $time_entries |
    Select-Object @SelectObjectArgs |
        Export-Csv -Path $PSScriptRoot\report.csv -notype

#Destroy the session
#TODO: I suspect this doesn't really kill the session
Invoke-RestMethod "$($config.toggl_api_uri)/sessions" -Method Delete -WebSession $toggl_api_session -Verbose > $null

# Open generated file
Invoke-Item -Path "$PSScriptRoot\report.csv"