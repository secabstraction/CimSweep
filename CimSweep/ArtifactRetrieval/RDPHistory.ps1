function Get-CSRDPHistory {
<#
.SYNOPSIS

Retrieves RDP session history.

Author: Jesse Davis (@secabstraction)
Props: Brandon Arvanaghi (@arvanaghi) 
PortedFrom: https://github.com/fireeye/SessionGopher/blob/master/SessionGopher.ps1
License: BSD 3-Clause

.DESCRIPTION

Get-CSRDPHistory retrieves RDP session history information stored in the registry.

.PARAMETER CimSession

Specifies the CIM session to use for this cmdlet. Enter a variable that contains the CIM session or a command that creates or gets the CIM session, such as the New-CimSession or Get-CimSession cmdlets. For more information, see about_CimSessions.

.EXAMPLE

Get-CSRDPHistory

.EXAMPLE

Get-CSRDPHistory -CimSession $CimSession

.OUTPUTS

CimSweep.RDPSession

Outputs objects consisting of relevant RDP session information.

#>

    [CmdletBinding()]
    [OutputType('CimSweep.RDPSession')]
    param (
        [Alias('Session')]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Management.Infrastructure.CimSession[]]
        $CimSession
    )
    
    begin {
        # If a CIM session is not provided, trick the function into thinking there is one.
        if (-not $PSBoundParameters['CimSession']) {
            $CimSession = ''
            $CIMSessionCount = 1
        } else {
            $CIMSessionCount = $CimSession.Count
        }

        $CurrentCIMSession = 0
    }

    process {
        foreach ($Session in $CimSession) {
            $ComputerName = $Session.ComputerName
            if (-not $Session.ComputerName) { $ComputerName = 'localhost' }

            # Display a progress activity for each CIM session
            Write-Progress -Id 1 -Activity 'CimSweep - RDP History sweep' -Status "($($CurrentCIMSession+1)/$($CIMSessionCount)) Current computer: $ComputerName" -PercentComplete (($CurrentCIMSession / $CIMSessionCount) * 100)
            $CurrentCIMSession++

            $CommonArgs = @{}

            if ($Session.Id) { $CommonArgs['CimSession'] = $Session }
            
            $UserSids = Get-HKUSID @CommonArgs
            
            foreach ($Sid in $UserSids) {

                $Parameters = @{
                    Hive = 'HKU'
                    SubKey = "$Sid\Software\Microsoft\Terminal Server Client\Servers"
                }
    
                Get-CSRegistryKey @Parameters @CommonArgs | Get-CSRegistryValue -ValueName UsernameHint | ForEach-Object {
                 
                    $ObjectProperties = [ordered] @{ 
                        PSTypeName = 'CimSweep.RDPSession'
                        UserSid = $Sid
                        UsernameHint = $_.ValueContent
                        Server = Split-Path -Leaf $_.SubKey
                    }

                    if ($_.PSComputerName) { $ObjectProperties['PSComputerName'] = $_.PSComputerName }
                    [PSCustomObject]$ObjectProperties
                }
            } 
        }
    }
    end {}
}

Export-ModuleMember -Function Get-CSRDPHistory