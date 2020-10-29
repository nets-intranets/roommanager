. "$PSScriptRoot\.helpers.ps1"
$context = (Init $MyInvocation  $false)

$counter = 0
foreach ($room in $ExchangeCurrentRooms.Values) {
    $counter++
    $percent = [int]($counter / $ExchangeCurrentRooms.Count * 100)

    Write-Progress -Activity "Syncronizing $($ExchangeCurrentRooms.Count) rooms with SharePoint list" -Status "$percent% Complete:" -PercentComplete $percent -CurrentOperation "Checking $($room.PrimarySmtpAddress)"
    $url = ($site + '/Lists/Rooms/items?$expand=fields&$filter=fields/Title eq ''' + $room.PrimarySmtpAddress + '''') 
    
    #Could have used a dictionary of rooms here, just like to excersise the list to ensure that indexes is in place
    $SharePointItem = Invoke-RestMethod $url -Method 'GET' -Headers $headers 
    
    if ($SharePointItem.value.length -eq 0) {
        Write-Progress -Activity "Syncronizing $($ExchangeCurrentRooms.Count) rooms with SharePoint list" -Status "$percent% Complete:" -PercentComplete $percent -CurrentOperation "Creating record for $($room.PrimarySmtpAddress)"

        
        $body =
        @"
{"fields":{
    "Title": "$($room.PrimarySmtpAddress)",
    "Display_x0020_Name": "$($room.DisplayName)",

  }
}
"@        

        $item = Invoke-RestMethod ($site + '/Lists/Rooms/items') -Method 'POST' -Headers $Headers  -Body $body
    }

}

Write-Progress -Completed  -Activity ""
Write-Output "Done importing"

