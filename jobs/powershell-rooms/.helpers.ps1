[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
function DotEnvConfigure($debug,$path) {
#    $path = $PSScriptRoot 
   $loop = $true
    
    do {
        $filename = "$path\.env"
        if ($debug) {
            write-output "Checking  $filename"
        }
        if (Test-Path $filename) {
            if ($debug) {
                write-output "Using $filename" 
            }
            $lines = Get-Content $filename
             
            foreach ($line in $lines) {
                    
                $nameValuePair = $line.split("=")
                if ($nameValuePair[0] -ne "") {
                    if ($debug) {
                        write-output "Setting >$($nameValuePair[0])<"
                    }
    
                    [System.Environment]::SetEnvironmentVariable($nameValuePair[0], $nameValuePair[1])
                }
            }
    
            $loop = $false
        }
        else {
            $lastBackslash = $path.LastIndexOf("\")
            if ($lastBackslash -lt 4) {
                $loop = $false
                if ($debug) {
                    write-output "Didn't find any .env file  "
                }
            }
            else {
                $path = $path.Substring(0, $lastBackslash)
            }
        }
    
    } while ($loop)
    
}
    

function GetAccessToken($client_id, $client_secret, $client_domain) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/x-www-form-urlencoded")
    $body = "grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret&scope=https%3A//graph.microsoft.com/.default"
    
    $response = Invoke-RestMethod "https://login.microsoftonline.com/$client_domain/oauth2/v2.0/token" -Method 'POST' -Headers $headers -body $body
    return $response.access_token
    
}
function ConnectExchange($username, $secret) {
    write-output "Connecting to Exchange Online"
    $code = ConvertTo-SecureString $secret -AsPlainText -Force
    $psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $code)
    
    
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $psCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking 
    return $Session
    
}
    
function CreateAlias($name) {
    return $name.ToLower().Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-").Replace(" ", "-")
}

function EnsurePath($path) {

    If (!(test-path $path)) {
        New-Item -ItemType Directory -Force -Path $path
    }
}

function RealErrorCount() {
    $c = 0
    foreach ($e in $Error) {
        $m = $e.ToString()
        if (!$m.Contains("__Invoke-ReadLineForEditorServices")) {
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


function LogToSharePoint($token, $site , $title, $status, $system, $subSystem, $reference, $Quantity, $details) {
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

    # write-output $body 
    #    Out-File -FilePath "$PSScriptRoot\error.json" -InputObject $body
    $url = ($site + '/Lists/Log/items/')
  
    $dummy = Invoke-RestMethod $url -Method 'POST' -Headers $myHeaders -Body $body 
    return $null -eq $dummy
}

function ReportErrors($token, $site) {
    if ($Error.Count -gt 0) {
        $errorMessages = ""
        foreach ($errorMessage in $Error) {
            if (($null -ne $errorMessage.InvocationInfo) -and ($errorMessage.InvocationInfo.ScriptLineNumber)) {
                $errorMessages += ("Line: " + $errorMessage.InvocationInfo.ScriptLineNumber + " "  )    
            }

            $errorMessages += $errorMessage.ToString() 
            $errorMessages += "`n"

        }

        LogToSharePoint $token $site "Error in PowerShell" "Error" "PowerShell"  $MyInvocation.MyCommand $null 0 $errorMessages
    }



    function ConnectExchange($username, $secret) {
        $code = ConvertTo-SecureString $secret -AsPlainText -Force
        $psCred = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $code)
    
    
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $psCred -Authentication Basic -AllowRedirection
        Import-PSSession $Session -DisableNameChecking 
        return $Session
    
    }
    
}


function FindSiteByUrl($token, $siteUrl) {
    $Xheaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $Xheaders.Add("Content-Type", "application/json")
    $Xheaders.Add("Prefer", "apiversion=2.1") ## Not compatibel when reading items from SharePointed fields 
    $Xheaders.Add("Authorization", "Bearer $token" )

    $url = 'https://graph.microsoft.com/v1.0/sites/?$top=1'
    $topItems = Invoke-RestMethod $url -Method 'GET' -Headers $Xheaders 
    if ($topItems.Length -eq 0) {
        Write-Warning "Cannot read sites from Office Graph - sure permissions are right?"
        exit
    }

    $siteUrl = $siteUrl.replace("sharepoint.com/","sharepoint.com:/")
    $siteUrl = $siteUrl.replace("https://","")
    

    $url = 'https://graph.microsoft.com/v1.0/sites/' + $siteUrl 

    $site = Invoke-RestMethod $url -Method 'GET' -Headers $Xheaders 
   

    return  ( "https://graph.microsoft.com/v1.0/sites/" + $site.id)
}


function SharePointRead($context, $path) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $($context.token)" )
    $url = $context.site + $path
    if ($context.verbose) {
        write-output "SharePointRead $url"
    }
    $result = Invoke-RestMethod ($url) -Method 'GET' -Headers $headers 
    return $result.value

}

function SharePointLookup($context, $path) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $($context.token)" )
    $url = $context.site + $path
    if ($context.verbose) {
        write-output "SharePointLookup $url"
    }
    $result = Invoke-RestMethod ($url) -Method 'GET' -Headers $headers 
    return $result

}

function Init ($invocation, $requireExchange) {
    $scriptName = $invocation.MyCommand.Name
    $path = Split-Path $invocation.MyCommand.Path
    
    
    DotEnvConfigure $true $path

    $homePath =  (Resolve-Path ($path+ "\..\..\..")).Path


    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
   
    $tenantDomain =  $ENV:AADDOMAIN
    if ($null -eq $tenantDomain) {
        Write-Warning "Missing domain suffix - is \$env:AADDOMAIN defined?"
        exit
    }

    
    EnsurePath "$homePath\logs"

    $logPath = "$homePath\logs\$tenantDomain"
    EnsurePath $logPath

    $timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd-hh-mm-ss")
 
    Start-Transcript -Path "$logPath\$scriptName-$timestamp.log"

    EnsurePath "$homePath\data"

    $dataPath = "$homePath\data" #TODO: Make data path tenant aware
    $dataPath = "$datapath\$tenantDomain" #TODO: Make data path tenant aware
    EnsurePath $dataPath


    $Error.Clear() 
 

    $dev = $env:ROOMMGRAPPDEBUG
    $token = GetAccessToken $env:ROOMMGRAPPCLIENT_ID $env:ROOMMGRAPPCLIENT_SECRET $env:ROOMMGRAPPCLIENT_DOMAIN
    $site = FindSiteByUrl $token $env:SITEURL
    
    if ($null -eq $site) {
        Write-Warning "Not able for find site - is \$env:SITEURL defined?"
        exit
    }
    $context = @{
        $logPath = "$logPath\$tenantDomain"
        domain =  $tenantDomain
        IsDev    = $dev
        site     = $site
        datapath = $dataPath
        logpath  = $logPath
        token    = $token
        session  = $session
        siteUrl  = $env:SITEURL
    }


    if ($requireExchange) {
        $errorCount = (RealErrorCount)
        $session = ConnectExchange $env:AADUSER $env:AADPASSWORD
        if ($errorCount -ne (RealErrorCount)){
            Write-Warning "Cannot connect to Exchange"
            Done $context
            exit
        }

    }
    return $context

}
    



function Done($context) {

    
    if (!$context.IsDev) {
        write-output "Closing sessions"
        get-pssession | Remove-PSSession
    }

    ReportErrors $context.token $context.site
    Stop-Transcript
}

