. "$PSScriptRoot\.helpers.ps1"

$context = (Init $MyInvocation.MyCommand.Name $true)
    
$title = "Updating existing Rooms from Masterdata list"
write-host $title

$rooms = Get-Content "$($context.datapath)\rooms-masterdata.json" | Out-String | ConvertFrom-Json


$counter = 0
foreach ($room in $rooms) {
    $counter++
    $percent = [int]($counter /$rooms.length * 100)
    Write-Progress -Activity "Updating $($rooms.length) rooms" -Status "$percent% Complete:" -PercentComplete $percent -CurrentOperation "Processing $($room.PrimarySmtpAddress)"

    Set-Place $room.primarySMTPAddress `
    -Phone $room.phone `
    -Street $room.street `
    -AudioDeviceName $room.audioDeviceName `
    -DisplayDeviceName $room.DisplayDeviceName `
    -Building $room.building `
    -Floor $room.floor `
    -State $room.state `
    -Label $room.state `
    -CountryOrRegion $room.countryOrRegion `
    -City $room.city `
    -VideoDeviceName $room.videoDeviceName `
    -FloorLabel $room.floorLabel `
    -Capacity $room.capacity `
    -IsWheelChairAccessible $room.isWheelChairAccessible `
    -GeoCoordinates $room.geoCoordinates 

    
#https://office365itpros.com/2019/08/28/populating-locations-outlook-places-service/
}
Write-Progress -Completed  -Activity "done"


write-host "Done $title"
Done $context

