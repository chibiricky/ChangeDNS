<#
.SYNOPSIS
    A script to change DNS settings in bulk on all PCs on a particular OU
.DESCRIPTION
    This script scans through all the interfaces on all the computers in the specified OU. Only the DNS settings of the interfaces that match the specified IP address prefix are changed.
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
    [Parameter(Mandatory=$true)] [string] $OU,
    [Parameter(Mandatory=$true)] [string] $LocalIPPrefix,
    [Parameter(Mandatory=$true)] [string[]] $NewDNS,
    [Switch] $DryRun
)

$OUTemp1 = $OU -split "[\\/]"
$OUTemp2 = $OUTemp1[($OUTemp1.Length-1)..1]
$OUPath = "OU=" + ($OUTemp2 -join ",OU=") + ",DC=" + (($OUTemp1[0] -split "\.") -join ",DC=")

$Changed = 0
$Unchanged = 0
$Offline = 0
$Errors = 0

$Computers = (Get-ADComputer -Filter * -SearchBase $OUPath | Select-Object Name).Name
""
foreach ($Computer in $Computers) {
    $NICFound = $false
    Write-Host "$Computer" -NoNewline
    if (Test-Connection $Computer -Quiet -Count 2) {
        Write-Host ":"
        $NICs = Get-WMIObject Win32_NetworkAdapterConfiguration -ComputerName $Computer
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
                    }
                    else {
                        try {
                            $Result = $NIC.SetDNSServerSearchOrder($NewDNS)
                            if ($Result.ReturnValue -in (0,1)) {
                                Write-Host "  New DNS servers:     " -NoNewline
                                Write-Host $($NewDNS -join ", ") -ForegroundColor Green
                                $Changed++
                            }
                            else {
                                Write-Host "  SetDNSServerSearchOrder terminated with error code $($Result.ReturnValue)" -ForegroundColor Red
                                Write-Host "  Refer to https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/setdnsserversearchorder-method-in-class-win32-networkadapterconfiguration for more information about the error code."
                                $Errors++
                                }
                        }
                        catch {
                            Write-Host "  ERROR: Unable to change the DNS servers on $Computer" -ForegroundColor Red
                            Write-Host $_
                            $Errors++
                        }
                    }
                }
                else {
                    Write-Host "  No change is required" -ForegroundColor Blue
                    $Unchanged++
                }
            }

        }
        if (!$NICFound) {
            Write-Host "  No network interface is found with IP prefix $LocalIPPrefix" -ForegroundColor Blue
            $Unchanged++
        }
    }
    else {
        Write-Host " is " -NoNewline
        Write-Host "offline" -ForegroundColor Red
        $Offline++
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