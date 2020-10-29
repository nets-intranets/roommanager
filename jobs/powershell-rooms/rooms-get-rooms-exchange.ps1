. "$PSScriptRoot\.helpers.ps1"

$context = (Init $MyInvocation $true)

$ExchangeCurrentRooms = @{}
    
$title = "Loading existing Rooms from Exchange"
write-output $title
$mbxs = get-mailbox -RecipientTypeDetails:RoomMailbox
$counter = 0
foreach ($mailbox in $mbxs) {
    $counter++
    $percent = [int]($counter /$mbxs.Count * 100)
    Write-Progress -Activity "Reading $($mbxs.Count) rooms and processing policies from Exchange" -Status "$percent% Complete:" -PercentComplete $percent -CurrentOperation "Reading Policy $($mailbox.PrimarySmtpAddress)"

    $policies = Get-CalendarProcessing $mailbox.PrimarySmtpAddress

    Write-Progress -Activity "Reading $($mbxs.Count) rooms and processing policies from Exchange" -Status "$percent% Complete:" -PercentComplete $percent -CurrentOperation "Reading Place $($mailbox.PrimarySmtpAddress)"
    $place = Get-Place $mailbox.PrimarySmtpAddress

    $ExchangeCurrentRooms.Add($mailbox.PrimarySmtpAddress,
        @{ mailbox   = $mailbox
            policies = $policies
            place = $place
        }
    )
}
Write-Progress -Completed  -Activity "done"


ConvertTo-Json -InputObject $ExchangeCurrentRooms -Depth 10 | Out-File "$($context.datapath)\rooms-exchange.json"
write-output "Done $title"
Done $context

