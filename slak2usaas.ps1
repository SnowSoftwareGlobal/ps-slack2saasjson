#region configuration
$saas_Instance = "https://slack.com" ## basic url of the api
$saas_outputfolder = "C:\Temp" ## folder to store the results - no trailing backspace please
$script:saas_logfile = $saas_outputfolder + "\servicenowquery.log"
$script:saas_exportfile = $saas_outputfolder + "\slack.json"

## authentification stuff
$saas_accesstoken = 'xoxp-totally secret access token.' ## use your own one.

## setting preferences for writing messages - can be removed when working
$InformationPreference = "Continue"
$ErrorActionPreference = "Continue"
$DebugPreference = "SilentlyContinue"
#endregion

#region scriptvaribles
## put stuff here what is used by the script, but does not need to be configured by the user
#$saas_data_subscriptions = @()
$saas_data_users = New-Object System.Collections.ArrayList
$saas_data_subscriptions = New-Object System.Collections.ArrayList
$saas_data_subscriptionlist = New-Object System.Collections.ArrayList

$saas_recordlimit = 200
$saas_date_today = (get-date).ToString('yyyy-MM-dd')

#different api endpoints to be used - filter whenever possible and the api allows it
$saas_uri_apitest = $saas_instance + '/api/api.test'
$saas_uri_userlist = $saas_Instance + "/api/users.list?token=" + $saas_accesstoken + "&limit=" + $saas_recordlimit
$saas_uri_teambillable = $saas_Instance + "/api/team.billableInfo?token=" + $saas_accesstoken + "&limit=" + $saas_recordlimit
#endregion

## handy logger based on https://stackoverflow.com/a/38738942
function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","DEBUG")]
    [String]
    $Level = "INFO",
    [Parameter(Mandatory=$True)]
    [string]
    $Message,
    [Parameter(Mandatory=$False)]
    [string]
    $Path = $script:saas_logfile ##"$env:USERPROFILE\log.txt"
    )
    function TS {(get-date).ToString('yyyy-MM-ddTHHmmss')}
    $messagedata = ("[$(TS)][$Level]: $Message")
    switch ($Level){
    'INFO' { Write-Information -MessageData $messagedata}
    'WARN' { Write-Warning -Message $messagedata}
    'ERROR' { Write-Error -Message $messagedata}
    'DEBUG'  { Write-Debug -Message $messagedata}
    default { Write-Output -MessageData $messagedata}
    }
    $Messagedata | Out-File $Path -Append
}

###############################
# Here The Script begins!!! ###
###############################

Write-Log "INFO" "Start Script" $script:saas_logfile

## create securestring and header
$headers = @{ “Accept” = “application/json” }

#region autorize if required
Write-Log "INFO" ("Start authorization part") $script:saas_logfile

## having the token does the job
Write-Log "INFO" ("Finished Authorization part") $script:saas_logfile
#endregion

#region query user
Write-Log 'INFO' ("Start Invoking Users") $script:saas_logfile
$saas_tempuri = $saas_uri_userlist + '?token=' + $saas_accesstoken
$saas_result = $null

    do {
       try { $saas_result = Invoke-RestMethod -Uri $saas_tempuri -Method Get -Headers $headers
        } catch {
            $saas_result_currentException = $_.Exception.Response.StatusCode
            Write-Log 'ERROR' ("Querying users had an Error, returning " + $saas_result_currentException +".") $script:saas_logfile
        }
        Foreach ($data_tem in ($saas_result.members | Select id,name,@{Name='email';Expression={$_.profile.email}},@{Name='first_name';Expression={$_.profile.first_name}},@{Name='last_name';Expression={$_.profile.last_name}},deleted,is_bot)) {
            if ($date_tem.is_bot -eq 'True' -or $data_tem.name -eq 'slackbot') {} else {
                $userthing = [pscustomobject]@{
                    'Id' =  $data_tem.id;
                    'Email' = $data_tem.email;
                    'UserName' = $data_tem.name;
                    'Name' = ($data_tem.first_name + ' ' + $data_tem.last_name);
                    'Status' = switch ($data_tem.deleted) { true {'Deleted'} default {'Active'}}
                    <# not available
                    'FirstActivity' = ([datetime]$data_tem.sys_created_on).tostring('s');
                    'LastActivity' = ([datetime]$data_tem.last_login_time).tostring('s');
                    'Registration' = ([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($data_tem.updated))).tostring('s')
                    #>
                }
            $saas_data_users.add($userthing) > $null
            }
        $saas_tempuri = $saas_uri_userlist + '?token=' + $saas_accesstoken + '&cursor=' + $saas_result.response_metadata.next_cursor
        }
    } while ($saas_result.response_metadata.next_cursor)

Write-Log 'INFO' ("Finished Invoking Users") $script:saas_logfile
#end region

#region query subscriptions
Write-Log 'INFO' ("Start Invoking subscriptions") $script:saas_logfile
## only one plan there, so no need here.
## create a fake one.

$saas_data_subscriptionlist.add([pscustomobject]@{
    'ID' = 'Slack User Subscription';
    'MeteringType' = 'PerAssignedUser';
    'Type' = '0';
    'UserCap' = 8462 ##current max
    })

Write-Log 'INFO' ("finished Invoking subscriptions") $script:saas_logfile
#endregion


#region build subscriptions and query assignments
Write-Log 'INFO' ("Start builing subscriptions") $script:saas_logfile
foreach ($data_subscription in $saas_data_subscriptionlist) {
    $saas_tempuri = $saas_uri_teambillable + '?token=' + $saas_accesstoken
    $saas_result = $null
    $saas_assignedusers = New-Object System.Collections.ArrayList
    do {
        try {
            $saas_result = Invoke-RestMethod -Headers $headers -Method Get -Uri $saas_tempuri}
        catch {
            $currentException = $_.Exception.Response.StatusCode
            Write-Log 'ERROR' ("Querying subscriptions had an Error, returning " + $currentException +".") $script:saas_logfile
        }
        Foreach ($saas_userid in ($saas_result.billable_info | gm | where membertype -eq NoteProperty).Name) {
            if ($saas_result.billable_info.($saas_userid).billing_active -eq 'True') {
                $saas_assignedusers.add($saas_userid)
            }
        }
        $saas_tempuri = $saas_uri_teambillable + '?token=' + $saas_accesstoken + '&cursor=' + $saas_result.response_metadata.next_cursor
    }
    while ($saas_result.response_metadata.next_cursor)

    $subscriptionthing = [pscustomobject]@{
        'ID' = $data_subscription.ID;
        'UserCap' = $data_subscription.UserCap;
        'AvailableSeats' =  ($data_subscription.UserCap - $saas_assignedusers.Count);
        'MeteringType' = "PerAssignedUser";
        'Type' = @{
            'ID' = $data_subscription.ID;
            'Name' = $data_subscription.ID
            };
        'UserIds'= $saas_assignedusers
        <#  not delivered
        'Created' = '0';
        '
        #>
    }
    $saas_data_subscriptions.add($subscriptionthing) > $null
}

Write-Log 'INFO' ("Finished builing subscriptions") $script:saas_logfile
#endregion


#region build json document
Write-Log 'INFO' ("start building json file") $script:saas_logfile
    $jdoc = @{}
    $jdoc.add('Subscriptions',$saas_data_subscriptions)
    $jdoc.add('Users',$saas_data_users)
    $jdoc | convertto-json -depth 5 > $script:saas_exportfile ## where to write the file to.
Write-Log 'INFO' ("Finished builing json file") $script:saas_logfile
#endregion