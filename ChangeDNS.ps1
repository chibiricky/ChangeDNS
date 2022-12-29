<#
.SYNOPSIS
    A script to change DNS settings in bulk on all PCs on a particular OU
.DESCRIPTION
    This script scans through all the interfaces on all the computers in the specified OU. Only the DNS settings of the interfaces with the specified IP address prefix are changed. A summary will be written to a log file in the temp folder.
.PARAMETER OU
    Required.
    Specify the path of the OU, e.g. "example.com\parentOU\childOU".
    Forward slashes or back slashes can be used.
.PARAMETER LocalIPPrefix
    Required.
    Specify the prefix of the IP address that the target network interface is using, e.g. "192.168.0.*".
.PARAMETER NewDNS
    Required.
    Specify the new DNS servers, separate with commas if there are more than one addresses, e.g. "192.168.1.1", "192.168.1.2".
.PARAMETER DryRun
    Switch. Optional.
    If this is on, no actual change will be made, but a preview of the changes will be shown.
.INPUTS
    None
.OUTPUTS
    None
.NOTES
    Author:         Ricky Chan
    Creation Date:  2022/12/28
  
.EXAMPLE
    PS> ChangeDNS -OU "example.com\FirstOU\SecondOU" -LocalIPPrefix "10.0.1.*" -NewDNS "10.0.0.1", "10.0.0.2"

.EXAMPLE
    PS> ChangeDNS -OU "example.com\FirstOU\SecondOU" -LocalIPPrefix "10.0.1.*" -NewDNS "10.0.0.1", "10.0.0.2" -DryRun
#>


Param(
    [string] $OU,
    [string] $LocalIPPrefix,
    [string[]] $NewDNS,
    [string] $PrevLog,
    [Switch] $DryRun
)

$Computers = @()

if ($PrevLog -eq "") {
    if ($OU -eq "" -or $LocalIPPrefix -eq "" -or $null -eq $NewDNS) {
        Write-Host "ERROR: Please specify all the following parameters since PrevLog is not defined: OU, LocalIPPrefix, NewDNS" -ForegroundColor Red
        return
    }
    $OUTemp1 = $OU -split "[\\/]"
    $OUTemp2 = $OUTemp1[($OUTemp1.Length-1)..1]
    $OUPath = "OU=" + ($OUTemp2 -join ",OU=") + ",DC=" + (($OUTemp1[0] -split "\.") -join ",DC=")
    $Computers = (Get-ADComputer -Filter * -SearchBase $OUPath | Select-Object Name).Name
}
else {
    $FileData = Get-Content $PrevLog
    $ReadLine = $false
    foreach ($Line in $FileData) {
        if ($ReadLine) {
            if ($Line -ne "Offline:" -and $Line -ne "Error:" -and $Line -ne "") {
                $Computers += $Line    
            }
            else {
                $ReadLine = $false
            }
        }
        else {
            if ($Line -like "LocalIPPrefix*") {
                $LocalIPPrefix = ($Line -split ":")[1]
            }
            elseif ($Line -like "NewDNS*") {
                $NewDNS = (($Line -split ":")[1] -split ",")
            }
            elseif ($Line -eq "Offline:" -or $Line -eq "Error:") {
                $ReadLine = $true
            }

        }
    }
}

$Changed = 0
$Unchanged = 0
$Offline = 0
$Errors = 0

$ChangedArray = New-Object -TypeName 'System.Collections.ArrayList';
$UnchangedArray = New-Object -TypeName 'System.Collections.ArrayList';
$OfflineArray = New-Object -TypeName 'System.Collections.ArrayList';
$ErrorArray = New-Object -TypeName 'System.Collections.ArrayList';

""
foreach ($Computer in $Computers) {
    $NICFound = $false
    Write-Host "$Computer" -NoNewline
    if (Test-Connection $Computer -Quiet -Count 2) {
        Write-Host ":"
        try {
            $NICs = Get-WMIObject Win32_NetworkAdapterConfiguration -ComputerName $Computer
        }
        catch {
            Write-Host "  ERROR: Unable to change the DNS servers on $Computer" -ForegroundColor Red
            Write-Host $_
            $Errors++
            $ErrorArray.Add($Computer) | Out-Null
            continue
        }
        foreach ($NIC in $NICs) {
            if ((@($NIC.IPAddress) -like $LocalIPPrefix).Count -gt 0) {
                $NICFound = $true
                $DNSServers = $NIC.DNSServerSearchOrder
                "  Current DNS servers: " + ($DNSServers -join ", ")
                $ChangeRequired = $false
                $DNSServers | ForEach-Object {
                    if ($NewDNS -notcontains $_) {
                        $ChangeRequired = $true
                    }
                }
                $NewDNS | ForEach-Object {
                    if ($DNSServers -notcontains $_) {
                        $ChangeRequired = $true
                    }
                }
                if ($ChangeRequired) {             
                    if ($DryRun) {
                        Write-Host "  New DNS servers:     " -NoNewline
                        Write-Host $($NewDNS -join ", ") -ForegroundColor Green
                        $Changed++
                        $ChangedArray.Add($Computer) | Out-Null
                    }
                    else {
                        try {
                            $Result = $NIC.SetDNSServerSearchOrder($NewDNS)
                            if ($Result.ReturnValue -in (0,1)) {
                                Write-Host "  New DNS servers:     " -NoNewline
                                Write-Host $($NewDNS -join ", ") -ForegroundColor Green
                                $Changed++
                                $ChangedArray.Add($Computer) | Out-Null
                            }
                            else {
                                Write-Host "  SetDNSServerSearchOrder terminated with error code $($Result.ReturnValue)" -ForegroundColor Red
                                Write-Host "  Refer to https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/setdnsserversearchorder-method-in-class-win32-networkadapterconfiguration for more information about the error code."
                                $Errors++
                                $ErrorArray.Add($Computer) | Out-Null
                                }
                        }
                        catch {
                            Write-Host "  ERROR: Unable to change the DNS servers on $Computer" -ForegroundColor Red
                            Write-Host $_
                            $Errors++
                            $ErrorArray.Add($Computer) | Out-Null
                        }
                    }
                }
                else {
                    Write-Host "  No change is required" -ForegroundColor Blue
                    $Unchanged++
                    $UnchangedArray.Add($Computer) | Out-Null
                }
            }

        }
        if (!$NICFound) {
            Write-Host "  No network interface is found with IP prefix $LocalIPPrefix" -ForegroundColor Blue
            $Unchanged++
            $UnchangedArray.Add($Computer) | Out-Null
        }
    }
    else {
        Write-Host " is " -NoNewline
        Write-Host "offline" -ForegroundColor Red
        $Offline++
        $OfflineArray.Add($Computer) | Out-Null
    }
    Write-Host ""
}
if ($DryRun) {
    Write-Host "Result:  Would have changed: " -NoNewline
} else {
    Write-Host "Result:  Changed: " -NoNewline
}
    
Write-Host $Changed -ForegroundColor Green -NoNewline
Write-Host "; Unchanged: " -NoNewline
Write-Host $Unchanged -ForegroundColor Blue -NoNewline
Write-Host "; Offline: " -NoNewline
Write-Host $Offline -ForegroundColor Red -NoNewline
Write-Host "; Error: " -NoNewline
Write-Host $Errors -ForegroundColor Red
""

Write-Host "Writing log file $env:TEMP\ChangeDNS_$(Get-Date -Format "yyyyMMddHHmmss").log... " -NoNewline
$Output = "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss")`r`n"
$Output += "LocalIPPrefix:$LocalIPPrefix`r`n"
$Output += "NewDNS:$($NewDNS -join ',')`r`n`r`n"
$FirstItem = $true
if ($ChangedArray.Count -gt 0) {
    $Output += "Changed:`r`n$($ChangedArray -join "`r`n")`r`n"
    $FirstItem = $false
}
if ($UnchangedArray.Count -gt 0) {
    if(!$FirstItem) {$Output += "`r`n"}
    $Output += "Unchanged:`r`n$($UnchangedArray -join "`r`n")`r`n"
    $FirstItem = $false
}
if ($OfflineArray.Count -gt 0) {
    if(!$FirstItem) {$Output += "`r`n"}
    $Output += "Offline:`r`n$($OfflineArray -join "`r`n")`r`n"
    $FirstItem = $false
}
if ($ErrorArray.Count -gt 0) {
    if(!$FirstItem) {$Output += "`r`n"}
    $Output += "Error:`r`n$($ErrorArray -join "`r`n")`r`n"
}
$Output | Out-File -FilePath "$env:TEMP\ChangeDNS_$(Get-Date -Format "yyyyMMddHHmmss").log" -NoNewline
Write-Host "Done`r`n"
