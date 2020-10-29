#$env:ALLUSERSPROFILE

write-host "Setting environment variables"
write-host "-----------------------------------"
$lines = Get-Content  "$PSScriptRoot\.env" -Encoding:UTF8

foreach ($line in $lines) {
    $pair = $line.split("=")
    Write-Host "$($pair[0])"
    [Environment]::SetEnvironmentVariable($pair[0], $pair[1], "Machine")
}

write-host "-----------------------------------"

write-host "REMEBER TO DELETE THE .env FILE NOW"


