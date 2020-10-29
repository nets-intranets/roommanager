. "$PSScriptRoot\.helpers.ps1"
<#

#>

$context = (Init $MyInvocation  $true)

Write-Host "Batch Job starting"

$dev = $true #$env:CAVAAPPDEBUG

$client_id = $env:CAVAAPPCLIENT_ID
$client_secret = $env:CAVAAPPCLIENT_SECRET
$client_domain = $env:CAVAAPPCLIENT_DOMAIN
$site = $env:CAVAAPPSITE


function CreateRoom ($key, $name, $location, $capacity, $country, $videodevice) {

    #  https://o365info.com/room-mailbox-powershell-commands/
    $yy = "none"
    $alias = "room-$key" 
    # Remove-Mailbox "room-$key@$domain" -Confirm:$false   ## !! Doesn't work on rooms
    $SmtpAddress = "room-$key@$domain"
    $mbx = Get-Mailbox  $alias -ErrorAction SilentlyContinue 
    if (!$mbx) {
        $yy = New-Mailbox  -Alias $alias -PrimarySmtpAddress $SmtpAddress -Name $name -DisplayName $name -Room -ResourceCapacity  $capacity $err -ErrorAction SilentlyContinue 
        if ($err) {
            return @{status = "Error"
            result      = $err
        }
        }
        else {
            return @{status = "Created"
                result      = $mbx
            }
        }
    }
    else {
        return @{status = "Existed"
            result      = $mbx
        }
    }
    
    # Start-Sleep -s 10

    #    $res2 = Set-CalendarProcessing $alias -AddOrganizerToSubject $True 

    #    $res3 = Add-DistributionGroupMember -Identity "rooms-$location" -Member $alias
    #    $res4 = Set-Place  $alias   -Capacity $capacity -CountryOrRegion $country    -VideoDeviceName $videodevice   -Building $location
    #    $res5 = Add-MailBoxPermission $alias -User "room-role-receptionist@$domain" -AccessRights FullAccess
    #    $res6 = Add-RecipientPermission $alias -Trustee "room-role-receptionist@$domain" -AccessRights SendAs -Confirm:$False 
    
}
function GetRoom ($key) {

    #  https://o365info.com/room-mailbox-powershell-commands/

    $alias = "room-$key" 
    
    Get-Mailbox  "room-$key@$domain"
    
}

function CreateAlias($name) {
    return $name.ToLower().Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-")
}
function CreateRoomlist($displayName, $countryCode) {
    $alias = CreateAlias $displayName
    New-DistributionGroup -Alias "rooms-$alias"  -PrimarySmtpAddress "rooms-$alias@$domain" -Name $displayName  -RoomList 
}



function DeleteRoomlist($alias) {
    Remove-DistributionGroup "rooms-$alias" -Confirm:$false
}

function ChangeRoomlist($alias, $displayName) {
    Set-DistributionGroup "rooms-$alias" -Name $displayName  -RoomList 
}

function AddRoomToRoomlist ($key, $location) {
    $alias = "room-$key" 
    Add-DistributionGroupMember -Identity "rooms-$location" -Member $alias
}

function SetMailtip ($key, $mailTip) {

    $alias = "room-$key" 
    Set-Mailbox $alias -MailTip $mailTip


}

function GetCalendarProcessing ($key) {
    $alias = "room-$key" 
    Get-CalendarProcessing "room-$key@$domain" | fl

}

function SetRoomPolicyRestrictedToPA ($key) {

    $alias = "room-$key" 
    Set-CalendarProcessing $alias -AllBookInPolicy: $False -BookInPolicy:"room-role-pa@$domain", "room-role-receptionist@$domain" -AddOrganizerToSubject $True
    Set-Mailbox $alias -MailTip "Room can only be reserved by PA's and receptionists"
}

function SetRoomPolicyOpen ($key) {

    $alias = "room-$key" 
    Set-CalendarProcessing $alias -AllBookInPolicy: $True
    Set-Mailbox $alias -MailTip ""
}



$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")
$body = "grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret&scope=https%3A//graph.microsoft.com/.default"

$response = Invoke-RestMethod "https://login.microsoftonline.com/$client_domain/oauth2/v2.0/token" -Method 'POST' -Headers $headers -body $body
$token = $response.access_token

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Accept", "application/json")
$headers.Add("Authorization", "Bearer $token" )

$result = Invoke-RestMethod ($site + '/Lists/Batch%20Jobs/items?$expand=fields&$filter=fields/Status eq null') -Method 'GET' -Headers $headers 
write-host $result.value.length "jobs in queue"


if ($result.value.length -eq 0) {
    write-host "No requests"
    return
}

$locks = Invoke-RestMethod ($site + '/Lists/Locks/items?$expand=fields&$top=1') -Method 'GET' -Headers $headers 

if ($locks.value.length -ne 0) {
    # $createdDate = $locks.value[0].fields.Created.replace("Z","")
    # $lockcreated =  [datetime]::ParseExact($createdDate,"yyyy-mm-dd hh:mm",$null)
    # $lockcreated1 =  [System.Management.ManagementDateTimeConverter]::ToDateTime($createdDate)
    
    # $now = Get-Date
    # $diff = $now - $lockcreated
    if (!$dev) { 
        
        write-host "Will not continue as the session is locked"
        write-host "Locked" $locks.value[0].fields.Created "by" $locks.value[0].createdBy.user.displayName
        return
    }
    else {
        write-host "DEVMODE: Ignoring lock"
    }
}
else {
    
    write-host "Creating lock"
    $lockInfo = "$($env:COMPUTERNAME) $($MyInvocation.InvocationName.replace("\","\\"))"

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Authorization", "Bearer $token" )
    $lockInfo = "{
    `n    `"fields`": {
    `n        `"Title`": `"$lockInfo`"
    `n    }
    `n}"
    $response = Invoke-RestMethod($site + '/Lists/Locks/items') -Method 'POST' -Headers $headers -Body  $lockInfo
}
$code = $env:AADPASSWORD
$username = $env:AADUSER 
$domain = $env:AADDOMAIN
 
$password = ConvertTo-SecureString $code -AsPlainText -Force
$psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)

if ($Session -eq $null) {
    Write-host "Connecting to Exchange Online"

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $psCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking #-CommandName New-MailContact
}else {
    Write-host "Reusing connection to Exchange Online"
}


foreach ($item in $result.value) {
    $body = "{
        `n    `"fields`": {
        `n        `"Status`": `"Running`"
        `n    }
        `n}"
        
    $url = ($site + '/Lists/Batch%20Jobs/items/' + $item.id)
    write-host "Setting status to Running" $url
    $response = Invoke-RestMethod $url -Method 'PATCH' -Headers $headers -Body $body
    
    $json = $item.fields.Payload 
    $json2 = $json.Replace(",`r`n}", "}")

    write-host $item.id $item.fields.Title # $item.fields.Status
    write-host "*************************************************************************************"
    # write-host $json2

    $requests = ConvertFrom-Json -InputObject  $json2
    
    $results = @()

    ForEach ($obj in  $requests) {
        $action = $obj.action

        write-host "action"    
        write-host "----------------------------------"    
        write-host $action
        switch ($action) {
            "CreateRoom" {
                $output = CreateRoom $obj.key $obj.name $obj.location $obj.capacity $obj.country $obj.videodevice  
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }
            "CreateRoomlist" {
                $output = CreateRoomlist $obj.displayName $obj.country 
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }

            "ChangeRoomlist" {
                $output = ChangeRoomlist $obj.alias $obj.displayName
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }

 
            "DeleteRoomlist" {
                $output = DeleteRoomlist $obj.alias
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }

            "AddRoomToRoomlist" {
                $output = AddRoomToRoomlist $obj.roomKey $obj.locationKey
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }
            "GetRoom" {
                $output = GetRoom $obj.key 
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                    upn            = $obj.key 
                }
                break
            }
            "GetCalendarProcessing" {
                $output = GetCalendarProcessing $obj.key 
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                    upn            = $obj.key 
                }
                break
            }

              
            "SetMailtip" {
                $output = SetMailtip $obj.roomKey $obj.mailTip
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }

            "SetRoomPolicyRestrictedToPA" {
                $output = SetRoomPolicyRestrictedToPA $obj.roomKey 
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }
            "SetRoomPolicyOpen" {
                $output = SetRoomPolicyOpen $obj.roomKey 
                $result = @{output = if ( $output) { $output } else { "" }
                    action         = $action
                }
                break
            }
            default {
                $result = @{
                    output = "Unknown"
                    error  = "Unknown action"
                    action = $action
                }
            }

        } 

        $results += $result
           
    }
    
    $resultJSON = ConvertTo-Json -InputObject $results
    write-host "----------------------------------"    
    write-host "result"    
    write-host $resultJSON 
    write-host "----------------------------------"    
    $body = "{
        `n    `"fields`": {
        `n        `"Status`": `"Done`"
        `n        `"Result`": $resultJSON
        `n    }
        `n}"

    write-host $body 
    $url = ($site + '/Lists/Batch%20Jobs/items/' + $item.id)
    write-host "Setting status to Done" $url
    $response = Invoke-RestMethod $url -Method 'PATCH' -Headers $headers -Body $body
}



if (!$dev -and $Session) {

    $locks = Invoke-RestMethod ($site + '/Lists/Locks/items?$expand=fields&$top=1') -Method 'GET' -Headers $headers 

    if ($locks.value.length -ne 0) {
        $url = ($site + '/Lists/Locks/items/' + $locks.value[0].id)
        write-host "UnLocking"    
        $response = Invoke-RestMethod $url -Method 'DELETE' -Headers $headers -Body $body
    
        write-host "UnLocked"
        
    }

    write-host "Closing session"
    Remove-PSSession $Session
    $Session = $null
  
}
Done $context

