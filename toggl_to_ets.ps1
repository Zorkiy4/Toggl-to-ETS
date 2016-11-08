param (
    [string]$since = [DateTime]::Today.AddDays(-1).ToString('yyyy-MM-dd'), 
    [string]$until = [DateTime]::Today.AddDays(-1).ToString('yyyy-MM-dd')
)

#Load configuration from file
. $PSScriptRoot\config.ps1

# Function accepts uername and password, returns a hash to be used as a
#  -Headers parameter for an Invoke-RestMethod or Invoke-WebRequest
function New-BasicAuthHeader {
  [cmdletbinding()]
  param (
    [Parameter(Mandatory=$true)] $Username,
    [Parameter(Mandatory=$true)] $Password
  )
  # Doing this step-by-step for illustration; no reason not to reduce it to fewer steps
  # In fact, you can uncomment the following line and remove all the following lines in the function
  # @{ Authorization = "Basic {0}" -f [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $Username,$Password))) }

  # Make "username:password" string
  $UserNameColonPassword = "{0}:{1}" -f $Username,$Password
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

    foreach($entry in $time_entries) {
 
        if($curr_date.Date -gt $entry.start.Date) {
            $curr_date = $entry.start
        }
        elseif ($curr_date.Date -ne $entry.start.Date) {
            Write-Host "Exiting with code 365: unexpected date order" 
            exit 365
        }

        $same_task = $sum_entries.Where({$_.description -eq $entry.description -and $_.start.Date -eq $entry.start.Date}, 'First')
         
        if($same_task) {
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
        [Parameter(Mandatory=$true, ValueFromPipeline = $true)][System.Object] $obj,
        [Parameter(Mandatory=$true)][System.String[]] $prop_names
    )

    foreach($prop in $prop_names) {
        $obj.$prop = [datetime]::ParseExact($obj.$prop, 'yyyy-MM-ddTHH:mm:sszzz', [System.Globalization.CultureInfo]::CurrentCulture)
    }

    #Return object
    $obj
}

function PrepareTo-ETS {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline = $true)][System.Object[]] $entries,
        [Parameter(Mandatory=$true)]$irregular_time_start,
        [Parameter(Mandatory=$true)]$irregular_time_end,
        [Parameter(Mandatory=$true)]$irregular_time_prefix
    )

    #Join entries with the same description
    $sum_entries = Summarize-Entries -time_entries $entries

    foreach ($entry in $sum_entries) {

        #Convert duration from milliseconds to hours
        $entry.dur = [Math]::Round($entry.dur/3600000, 1)
        $entry.tags = $entry.tags[0]

        #Set 'none' tag if nothing was set
        if($entry.tags -eq $null) {
            $entry.tags = 'none'
        }

        #Prepend tag with "Irregular hours" if needed
        if(!($entry.tags -contains 'Irregular hours') -and ($entry.start.Hour -ge $irregular_time_start -or $entry.start.Hour -lt $irregular_time_end -or $entry.start.DayofWeek -eq 0 -or $entry.start.DayofWeek -eq 6)) {
            $entry.tags = $irregular_time_prefix + $entry.tags
        }
    }

    $sum_entries
}


#####################################################################
#                           Main code 
#####################################################################

$Headers = New-BasicAuthHeader -Username $api_token -Password "api_token"

#Authentication and session creation
Invoke-RestMethod "https://www.toggl.com/api/v8/sessions" -Method Post -Headers $Headers -SessionVariable toggl_api_session -Verbose > $null

#Get the client data by client name
$client = Invoke-RestMethod "https://www.toggl.com/api/v8/me?with_related_data=true" -Method Get -ContentType "application/json" -WebSession $toggl_api_session -Verbose | 
            Select-Object -ExpandProperty data | 
            Select-Object -ExpandProperty clients |
            Where-Object {$_.name -eq $client_name}

#Get projects related to the given client
$uri = "https://www.toggl.com/api/v8/clients/" + $client.id + "/projects"

$projects = Invoke-RestMethod $uri -Method Get -ContentType "application/json" -WebSession $toggl_api_session -Verbose

#Convert array of objects to comma separated string on IDs
$project_ids = '';
foreach ($project in $projects) {
    if($project_ids -ne '') {$project_ids += ','}
    $project_ids += $project.id
}

#Load time entries
#Results might be splitted to several pages. So we need a loop to load them all.

$page = 0
$loaded = 0
$time_entries = $null

do {
    $uri = "https://toggl.com/reports/api/v2/details?workspace_id=" + $client.wid +"&project_ids=" + $project_ids + "&since=" + $since + "&until=" + $until + "&user_agent=api_test&page=" + $page
    $result = Invoke-RestMethod $uri -Method Get -ContentType "application/json" -WebSession $toggl_api_session -Verbose
    $time_entries +=  $result | Select -ExpandProperty data | ForEach-Object -process {string_to_datetime -obj $_ -prop_names @("start", "end")}
    $loaded += $result.per_page
    $page++

} while ($loaded -lt $result.total_count)

PrepareTo-ETS -entries $time_entries -irregular_time_start $irregular_time_start -irregular_time_end $irregular_time_end -irregular_time_prefix $irregular_time_prefix | 
    Select -Property @{Name="Project-Task"; Expression={$_.project + "." + $_.tags}}, 
        @{Name="Effort"; Expression={$_.dur}}, 
        @{Name="Description"; Expression={$_.description}}, 
        @{Name="Started Date"; Expression={$_.start.ToString('d')}},
        @{Name="Completion Date"; Expression={$_.end.ToString('d')}} |
    Export-Csv -Path $PSScriptRoot\report.csv -notype

#Destroy the session
#TODO: I suspect this doesn't really kill the session
Invoke-RestMethod "https://www.toggl.com/api/v8/sessions" -Method Delete -WebSession $toggl_api_session -Verbose > $null