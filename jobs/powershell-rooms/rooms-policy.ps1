. "$PSScriptRoot\.helpers.ps1"
<#



#>

. "$PSScriptRoot\.helpers.ps1"
EnsurePath  "$PSScriptRoot\..\data"
Out-File "$PSScriptRoot\..\data\text.txt" -InputObject "test"

$Error.Clear() 

Write-Host "Policy Job starting"

$dev = $env:ROOMMGRAPPDEBUG

$client_id = $env:ROOMMGRAPPCLIENT_ID
$client_secret = $env:ROOMMGRAPPCLIENT_SECRET
$client_domain = $env:ROOMMGRAPPCLIENT_DOMAIN
$site = $env:ROOMMGRAPPSITE

if ($null -eq $site){
    Write-Warning "Missing env:ROOMMGRAPPSITE"
    Write-Warning "Exiting"
    exit
}

$CHECK_POLICY = $true
$PREFIX_ROOMPOLICY = "room-policy-"

$SUFFIX_GROUPALIAS = ""
$SUFFIX_GROUPNAME = ""
$POLICY_TEXT_PREFIX = "Room Policy " # Observe the trailing space


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


Write-host "Connecting to Exchange Online"
$code = $env:AADPASSWORD
$username = $env:AADUSER 
$domain = $env:AADDOMAIN
 
$password = ConvertTo-SecureString $code -AsPlainText -Force
$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)

if ($Session -eq $null) {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $psCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking #-CommandName New-MailContact
}

$ExchangeCurrentRoomGroups = @{}
$ExchangeCurrentRooms = @{}
    
function CreateAlias($name) {
    return $name.ToLower().Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-")
}

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

function isMember($members, $roomSmtpAddress) {
    $found = $false
    foreach ($member in $members) {
        if ($members.PrimarySmtpAddress -eq $roomSmtpAddress) {
            $found = $true
        }
    }
    return $found
}

function LoadCurrentData() {
    write-host "Loading existing Room Lists from Exchange"
    $dls = get-distributiongroup -RecipientTypeDetails RoomList 
    foreach ($dl in $dls) {
        $members = get-distributiongroupmember $dl.identity
        $ExchangeCurrentRoomGroups.Add($dl.PrimarySmtpAddress, $members)
    
    }

    $mbxs = get-mailbox -RecipientTypeDetails:RoomMailbox
    foreach ($mbx in $mbxs) {
        $ExchangeCurrentRooms.Add($mbx.PrimarySmtpAddress, $mbx)
    }
    write-host "Done Loading existing Room Lists from Exchange"
}


function EnsurePolicies() {
   

    $url = $site + '/Lists/Rooms/items?$expand=fields&$top=5000' 
    

    $result = Invoke-RestMethod ($url) -Method 'GET' -Headers $headers 
    write-host $result.value.length "rooms"

    $number = 0
    foreach ($room in $result.value) {
        $number += 1
        write-host $number "Checking" $room.fields.Title
        $odata = $site + '/Lists/Room%20Policies/items/' + $room.fields.PolicyLookupId + '?$expand=fields'
        $policy = Invoke-RestMethod ($odata) -Method 'GET' -Headers $headers 

        $roomSmtpAddress = $room.fields.Title


        
        $existingRoom = $ExchangeCurrentRooms[$room.fields.Title] 

        if (!$existingRoom ) {
            write-host "Creating new room mailbox"
            $yy = New-Mailbox -PrimarySmtpAddress $roomSmtpAddress -Name  $room.fields.Display_x0020_Name -DisplayName $room.fields.Display_x0020_Name -Room 
            LogToSharePoint "Created Room $roomSmtpAddress " "Info" "Exchange" "New-Mailbox " $roomSmtpAddress  1


        }

        if ( $room.fields.Display_x0020_Name -ne $existingRoom.DisplayName) {
            write-host "Updating DisplayName" $roomSmtpAddress $room.fields.Display_x0020_Name
            Set-Mailbox  $roomSmtpAddress -DisplayName  $room.fields.Display_x0020_Name

        }


        if ($policy.fields.RestrictBooking) {
            $policyAlias = $PREFIX_ROOMPOLICY + (CreateAlias $policy.fields.Title) + $SUFFIX_GROUPALIAS
            
            if (!$policy.fields.Mailtip) {
                Set-Mailbox $roomSmtpAddress -MailTip "Booking restricted to members of $($policy.fields.Title)"
            }
            else {
                Set-Mailbox $roomSmtpAddress -MailTip ($policy.fields.Mailtip + " Booking restricted to members of $($policy.fields.Title)")
            }
            Set-CalendarProcessing $roomSmtpAddress -AllBookInPolicy:$false -BookInPolicy:"$policyAlias@$domain"
        }
        


        if ($CHECK_POLICY) {
            # Baseline
            $currentProcessingPolicy = Get-CalendarProcessing $room.fields.Title 

            # Set to $true is changes has been applied
            $needUpdate = $false

            if (!$currentProcessingPolicy.DeleteComments) {
                $needUpdate = $true
                Set-CalendarProcessing  $room.fields.Title  -DeleteComments:$false
            }




            # Result
            $currentProcessingPolicy = Get-CalendarProcessing $room.fields.Title 


            if ($needUpdate) {


                $statustext = ""

                $statustext += "DeleteComments;" + $currentProcessingPolicy.DeleteComments + "`r`n"
                $statustext += "AddOrganizerToSubject;" + $currentProcessingPolicy.AddOrganizerToSubject + "`r`n"


                $statustext += "AllBookInPolicy;" + $currentProcessingPolicy.AllBookInPolicy + "`r`n"
           
                $statustext += "AdditionalResponse;" + $currentProcessingPolicy.AdditionalResponse + "`r`n"
                $statustext += "AdditionalResponse;" + $currentProcessingPolicy.AdditionalResponse + "`r`n"
                $statustext += "BookingWindowInDays;" + $currentProcessingPolicy.BookingWindowInDays + "`r`n"
                $statustext += "BookInPolicy;" + $currentProcessingPolicy.BookInPolicy + "`r`n"
     
    


                $resultJSON = ConvertTo-Json -InputObject $statustext
                # write-host "----------------------------------"    
                # write-host "result"    
                # write-host $resultJSON 
                # write-host "----------------------------------"    
                $body = "{
            `n    `"fields`": {
            `n        `"Applied_x0020_Policies`": $resultJSON
            `n    }
            `n}"
    
                # write-host $body 
                $url = ($site + '/Lists/Rooms/items/' + $room.id)
      
                $response = Invoke-RestMethod $url -Method 'PATCH' -Headers $headers -Body $body
            
            } 
        }
    
    }
}

function EnsurePolicyGroups() {
   
    write-host "Checking Policy Groups" 
    $odata = $site + '/Lists/Room%20Policies/items?$expand=fields'
    $policies = Invoke-RestMethod ($odata) -Method 'GET' -Headers $headers 

    foreach ($policy in $policies.value) {
        # I only create distributionlist if need (but don't delete existing )
        if ($policy.fields.RestrictBooking) {
            $policyAlias = $PREFIX_ROOMPOLICY + (CreateAlias $policy.fields.Title) 
            $policyGroup = Get-DistributionGroup $policyAlias  -ErrorAction SilentlyContinue 
            $policySmtpAddress = "$policyAlias@$domain"
            $displayName = $policy.fields.Title
            if (!$policyGroup) {

                write-host "Creating Policy Group "  $policySmtpAddress
                $errorBeforeCount = RealErrorCount
                New-DistributionGroup -Alias $policyAlias  -PrimarySmtpAddress $policySmtpAddress -Name $displayName -DisplayName ($POLICY_TEXT_PREFIX + $displayName)  -CopyOwnerToMember -ManagedBy $policy.fields.Owner1_x0020_Email 
                if ($errorBeforeCount -ne (RealErrorCount)) {
                    LogToSharePoint "Error creating Room Policy $policySmtpAddress" "Error" "Exchange" "New-DistributionGroup" $policySmtpAddress 1 (LastError)
                }
                else {
                    LogToSharePoint "Created Room Policy $policySmtpAddress" "OK" "Exchange" "New-DistributionGroup" $policySmtpAddress 1 ""
                }
                

            }
            if ($policy.fields.Owner1_x0020_Email) {
                Add-DistributionGroupMember $policySmtpAddress -Member $policy.fields.Owner1_x0020_Email  -ErrorAction SilentlyContinue 
            }
            if ($policy.fields.Owner2_x0020_Email) {
                Add-DistributionGroupMember $policySmtpAddress -Member $policy.fields.Owner2_x0020_Email  -ErrorAction SilentlyContinue 
            }
        }
    }

}



function RoomLists() {

   
    $roomSites = Invoke-RestMethod  ($site + '/Lists/Room%20Sites/items?$expand=fields&$top=5000') -Method 'GET' -Headers $headers 
    write-host $roomSites.value.length "sites"


    $roomSitesLookup = @{}
    $roomLists = @{} 
    
   
    foreach ($roomSite in $roomSites.value) {
        $roomSitesLookup.Add($roomSite.fields.Id.ToString(), $roomSite)
        $alias = createAlias $roomSite.fields.Title
        $roomGroupSmtpAddress = "rooms-$alias@$domain"

        # If it doesn't exists
        if (!$ExchangeCurrentRoomGroups.ContainsKey($roomGroupSmtpAddress)) {
            write-host "Creating RoomList for Site"
            New-DistributionGroup -PrimarySmtpAddress $roomGroupSmtpAddress -Name $roomSite.fields.Title  -RoomList 
        }
       
    }
    
    $rooms = Invoke-RestMethod ($site + '/Lists/Rooms/items?$expand=fields&$top=5000') -Method 'GET' -Headers $headers 
    write-host $result.value.length "rooms"


    foreach ($room in $rooms.value) {

        if ($room.fields.SiteLookupId) {
            $alias = (createAlias $roomSitesLookup[$room.fields.SiteLookupId].fields.Title)
            $roomGroupSmtpAddress = "rooms-$alias@$domain"
            $dl = $ExchangeCurrentRoomGroups[$roomGroupSmtpAddress]
            if ($dl) {
                if (!(isMember $dl $room.fields.Title)) {
                    Write-Host "Site - Adding $($room.fields.Title) to $roomGroupSmtpAddress"
                    $res = Add-DistributionGroupMember -Identity $dl.Identity -Member $room.fields.Title -ErrorAction SilentlyContinue
                }
            }
        }

        
    }
    


}


LoadCurrentData
RoomLists
EnsurePolicyGroups
EnsurePolicies


Write-Host "Done syncing"

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



