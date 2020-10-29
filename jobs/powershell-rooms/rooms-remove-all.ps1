. "$PSScriptRoot\.helpers.ps1"
<#



#>

Write-Host "Removing all rooms and distribution room lists"

$dev = $env:ROOMMGRAPPDEBUG

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

$dls = get-distributiongroup -RecipientTypeDetails RoomList 
foreach ($dl in $dls) {
    Remove-DistributionGroup $dl.PrimarySmtpAddress -Confirm:$false
}

$mbxs = get-mailbox -RecipientTypeDetails:RoomMailbox
foreach ($mbx in $mbxs) {
    Remove-Mailbox $mbx.PrimarySmtpAddress -Confirm:$false
}

if (!$dev -and $Session) {
    write-host "Closing session"cls
    Remove-PSSession $Session
    $Session = $null
  
}
