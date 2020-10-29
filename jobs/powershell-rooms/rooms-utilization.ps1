. "$PSScriptRoot\.helpers.ps1"
<#



#>



$Error.Clear() 

Write-Host "Room Utilization Job starting"

$dev = $env:ROOMMGRAPPDEBUG

$client_id = $env:ROOMMGRAPPCLIENT_ID
$client_secret = $env:ROOMMGRAPPCLIENT_SECRET
$client_domain = $env:ROOMMGRAPPCLIENT_DOMAIN
$site = $env:ROOMMGRAPPSITE



$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")
$body = "grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret&scope=https%3A//graph.microsoft.com/.default"

$response = Invoke-RestMethod "https://login.microsoftonline.com/$client_domain/oauth2/v2.0/token" -Method 'POST' -Headers $headers -body $body
$token = $response.access_token


function LogToSharePoint($title, $status, $system, $subSystem, $reference, $Quantity, $details) {
    $myHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $myHeaders.Add("Content-Type", "application/json")
    $myHeaders.Add("Accept", "application/json")
    $myHeaders.Add("Authorization", "Bearer $token" )
    $hostName = $env:COMPUTERNAME
    $details = $details -replace """", "\"""
    $body = "{
        `n    `"fields`": {
        `n        `"Title`": `"$title`",
        `n        `"Host`": `"$hostName`",
        `n        `"Status`": `"$status`",
        `n        `"System`": `"$system`",
        `n        `"SubSystem`": `"$subSystem`",
        `n        `"SystemReference`":`"$reference`",
        `n        `"Quantity`": $Quantity,
        `n        `"Details`": `"$details`"
        `n    }
        `n}"

    # write-host $body 
    #    Out-File -FilePath "$PSScriptRoot\error.json" -InputObject $body
    $url = ($site + '/Lists/Log/items/')
 
    $response = Invoke-RestMethod $url -Method 'POST' -Headers $myHeaders -Body $body
        
}




$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Accept", "application/json")
$headers.Add("Authorization", "Bearer $token" )



function GetFreeBuzy($upn, $date) {
    
    $day = $date.Day
    $month= $date.Month
    $year = $date.Year
    $dayOfWeek= $date.DayOfWeek

    $midNight = Get-Date -Year $year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0
    $midNightNextDay = $midNight.AddDays(1)
    $from =  ($midNight.ToString("yyyy-MM-dd") + "T00:00:00.0000000Z")
    $to =  ($midNightNextDay.ToString("yyyy-MM-dd") + "T00:00:00.0000000Z")

 
    $body = @"
    {
    	"EndTime": {
    		"dateTime": "$to",
    		"timeZone": "UTC"
    	},
    	"Schedules": ["$upn"],
    	"StartTime": {
    		"dateTime": "$from",
    		"timeZone": "UTC"
    	},
    	"availabilityViewInterval": "15"
    }
"@
        
    $response = Invoke-RestMethod "https://graph.microsoft.com/v1.0/users/$upn/calendar/getSchedule" -Method 'POST' -Headers $headers -Body $body
    $availabilityView = $response.value[0].availabilityView

    $timeslots = ""
    for ($i = 0; $i -lt 24; $i++) {
        $thisHour = $availabilityView.substring($i*4,4)
        $utilization = 0
        if ($thisHour[0] -ne "0"){
            $utilization += 0.25
        }else {
            
        }
        if ($thisHour[1] -ne "0"){
            $utilization += 0.25
        }else {
            
        }
        if ($thisHour[2] -ne "0"){
            $utilization += 0.25
        }else {
            
        }
        if ($thisHour[3] -ne "0"){
            $utilization += 0.25
        }else {
            
        }

        $timeslots += ("""Hour" +  $i + """:"+ $utilization + ",`n")
        #write-host "Hour $i Utilizaion $utilization"
    }
    Write-Host $timeslots
    # $response | ConvertTo-Json -Depth 10
}

GetFreeBuzy "niels@jumpto365.com" (Get-Date).ToUniversalTime()

function RealErrorCount() {
    $c = 0
    foreach ($e in $Error) {
        $m = $e.ToString()
        if (!$m.StartsWith("The term '__Invoke-ReadLineForEditorServices'")) {
            $c++
        }
    }
    return $c 
}
function LastError() {
    $m = ""
    foreach ($e in $Error) {
        $m += ($e.ToString().substring(0, 200) + "`n")

    }
    return $m    
}












if (!$dev -and $Session) {
    write-host "Closing session"cls
    Remove-PSSession $Session
    $Session = $null
  
}

if ($Error.Count -gt 0) {
    $errorMessages = ""
    foreach ($errorMessage in $Error) {
        if (($null -ne $errorMessage.InvocationInfo) -and ($errorMessage.InvocationInfo.ScriptLineNumber)) {
            $errorMessages += ("Line: " + $errorMessage.InvocationInfo.ScriptLineNumber + " "  )    
        }

        $errorMessages += $errorMessage.ToString() 
        $errorMessages += "`n"

    }

    LogToSharePoint "Error in PowerShell" "Error" "PowerShell"  $MyInvocation.MyCommand $null 0 $errorMessages
}
