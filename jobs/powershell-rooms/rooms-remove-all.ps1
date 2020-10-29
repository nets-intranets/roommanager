. "$PSScriptRoot\.helpers.ps1"
<#



#>

write-output "Removing all rooms and distribution room lists"

$dev = $env:ROOMMGRAPPDEBUG

write-output "Connecting to Exchange Online"
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
    write-output "Closing session"
    Remove-PSSession $Session
    $Session = $null
  
}
