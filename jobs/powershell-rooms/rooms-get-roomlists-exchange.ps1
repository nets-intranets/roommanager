. "$PSScriptRoot\.helpers.ps1"

$context = (Init $MyInvocation  $true)

$ExchangeCurrentRoomGroups = @{}
    
$title = "Loading existing Room Lists from Exchange"

write-output $title

$dls = get-distributiongroup -RecipientTypeDetails RoomList 
$counter = 0

foreach ($dl in $dls) {
    $counter++
    $percent = [int]($counter /$dls.Count * 100)
    Write-Progress -Activity "Reading $($dls.Count) room list and members from Exchange" -Status "$percent% Complete:" -PercentComplete $percent -CurrentOperation "Reading Members $($dl.PrimarySmtpAddress)"

    $members = get-distributiongroupmember $dl.identity
    $ExchangeCurrentRoomGroups.Add($dl.PrimarySmtpAddress,
        @{ list     = $dl
            members = $members

        }
    )
}
Write-Progress -Completed  -Activity "done"


ConvertTo-Json -InputObject $ExchangeCurrentRoomGroups -Depth 10 | Out-File "$($context.datapath)\room-lists-exchange.json"
write-output "Done $title"
Done $context

