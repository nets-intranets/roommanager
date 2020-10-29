. "$PSScriptRoot\.helpers.ps1"

$context = (Init $MyInvocation  $false)

$SharePointRooms = @{}    
$title = "Loading existing Room Lists from SharePoint"

write-host $title

$rooms = (SharePointRead  $context '/Lists/Rooms/items?$expand=fields&$top=5000')
$counter = 0

foreach ($roomItem in $rooms) {
    $counter++
    $percent = [int]($counter / $rooms.length * 100)
    $room = $roomItem.fields
     
     $policy = $null
     $site = $null
     $building = $null
     $location = $null
     $country = $null

    Write-Progress -Activity "Reading $($rooms.length) room list and members from SharePoint" -Status "$percent% Complete:" -PercentComplete $percent  -CurrentOperation "Room  $($room.Title)"

    if ($room.SiteLookupId) {
        $siteItem= (SharePointLookup  $context "/Lists/Room%20Sites/items/$($room.SiteLookupId)") 
        $site = $siteItem.fields

    }
    if ($room.PolicyLookupId) {
        $policyItem = (SharePointLookup  $context "/Lists/Room%20Policies/items/$($room.PolicyLookupId)") 
        $policy = $policyItem.fields
    
    }
    if ($room.BuildingLookupId) {
        $buildingItem = (SharePointLookup  $context "/Lists/Buildings/items/$($room.BuildingLookupId)") 
        $building = $buildingItem.fields
        if ($building.LocationLookupId) {
            $locationItem = (SharePointLookup  $context "/Lists/Locations/items/$($building.LocationLookupId)") 
            $location = $locationItem.fields
            if ($location.countryLookupId){
                $countryItem = (SharePointLookup  $context "/Lists/Countries/items/$($location.countryLookupId)") 
                $country = $countryItem.fields
    
            }
        }

    }
 
    $SharePointRooms.Add($room.Title,
        @{ room    = $room
            policy = $policy
            site   = $site
            building = $building
            location = $location
            country = $country

        }
    )
}
Write-Progress -Completed  -Activity "done"

ConvertTo-Json -InputObject $SharePointRooms -Depth 10 | Out-File "$($context.datapath)\rooms-sharepoint.json" 
write-host "Done $title"
Done $context

