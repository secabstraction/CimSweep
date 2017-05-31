﻿function Get-CSUserAssist {
<#
.SYNOPSIS

Retrieves and parses user assist entries.

Author: Jesse Davis (@secabstraction)
License: BSD 3-Clause

.DESCRIPTION

Get-CSUserAssist retrieves and parses user assist entry information stored in the registry.

.PARAMETER CimSession

Specifies the CIM session to use for this cmdlet. Enter a variable that contains the CIM session or a command that creates or gets the CIM session, such as the New-CimSession or Get-CimSession cmdlets. For more information, see about_CimSessions.

.EXAMPLE

Get-CSUserAssist

.EXAMPLE

Get-CSUserAssist -CimSession $CimSession

.OUTPUTS

CimSweep.UserAssistEntry

Outputs objects consisting of relevant user assist information. Note: the LastExecutedTime of this object is a UTC datetime string in Round-trip format.

#>

    [CmdletBinding()]
    [OutputType('CimSweep.UserAssistEntry')]
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

        # https://msdn.microsoft.com/en-us/library/bb882665.aspx
        $KnownFolderMapping = @{
            '{F38BF404-1D43-42F2-9305-67DE0B28FC23}' = 'Windows'
            '{18989B1D-99B5-455B-841C-AB7C74E4DDFC}' = 'Videos'
            '{F3CE0F7C-4901-4ACC-8648-D5D44B04EF8F}' = 'UsersFiles'
            '{0762D272-C50A-4BB0-A382-697DCD729B80}' = 'UserProfiles'
            '{5B3749AD-B49F-49C1-83EB-15370FBD4882}' = 'TreeProperties'
            '{A63293E8-664E-48DB-A079-DF759E0509F7}' = 'Templates'
            '{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}' = 'SystemX86'
            '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}' = 'System'
            '{0F214138-B1D3-4A90-BBA9-27CBC0C5389A}' = 'SyncSetup'
            '{289A9A43-BE44-4057-A41B-587A76D7E7F9}' = 'SyncResults'
            '{43668BF8-C14E-49B2-97C9-747784D784B7}' = 'SyncManager'
            '{B97D20BB-F46A-4C97-BA10-5E3608430854}' = 'Startup'
            '{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}' = 'StartMenu'
            '{A75D362E-50FC-4FB7-AC2C-A8BEAA314493}' = 'SidebarParts'
            '{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}' = 'SidebarDefaultParts'
            '{8983036C-27C0-404B-8F08-102D10DCFD74}' = 'SendTo'
            '{190337D1-B8CA-4121-A639-6D472D16972A}' = 'SearchHome'
            '{98EC0E18-2098-4D44-8644-66979315A281}' = 'SEARCH_MAPI'
            '{EE32E446-31CA-4ABA-814F-A5EBD2FD6D5E}' = 'SEARCH_CSC'
            '{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}' = 'SavedSearches'
            '{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}' = 'SavedGames'
            '{859EAD94-2E85-48AD-A71A-0969CB56A6CD}' = 'SampleVideos'
            '{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}' = 'SamplePlaylists'
            '{C4900540-2379-4C75-844B-64E6FAF8716B}' = 'SamplePictures'
            '{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}' = 'SampleMusic'
            '{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}' = 'RoamingAppData'
            '{8AD10C31-2ADB-4296-A8F7-E4701232C972}' = 'ResourceDir'
            '{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}' = 'RecycleBin'
            '{BD85E001-112E-431E-983B-7B15AC09FFF1}' = 'RecordedTV'
            '{AE50C081-EBD2-438A-8655-8A092E34987A}' = 'Recent'
            '{52A4F021-7B75-48A9-9F6B-4B87A210BC8F}' = 'QuickLaunch'
            '{2400183A-6185-49FB-A2D8-4A392A602BA3}' = 'PublicVideos'
            '{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}' = 'PublicPictures'
            '{3214FAB5-9757-4298-BB61-92A9DEAA44FF}' = 'PublicMusic'
            '{DEBF2536-E1A8-4C59-B6A2-414586476AEA}' = 'PublicGameTasks'
            '{3D644C9B-1FB8-4F30-9B45-F670235F79C0}' = 'PublicDownloads'
            '{ED4824AF-DCE4-45A8-81E2-FC7965083634}' = 'PublicDocuments'
            '{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}' = 'PublicDesktop'
            '{DFDF76A2-C82A-4D63-906A-5644AC457385}' = 'Public'
            '{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}' = 'Programs'
            '{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}' = 'ProgramFilesX86'
            '{6D809377-6AF0-444B-8957-A3773F02200E}' = 'ProgramFilesX64'
            '{DE974D24-D9C6-4D3E-BF91-F4455120B917}' = 'ProgramFilesCommonX86'
            '{6365D5A7-0F0D-45E5-87F6-0DA56B6A4F7D}' = 'ProgramFilesCommonX64'
            '{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}' = 'ProgramFilesCommon'
            '{905E63B6-C1BF-494E-B29C-65B732D3D21A}' = 'ProgramFiles'
            '{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}' = 'ProgramData'
            '{5E6C858F-0E22-4760-9AFE-EA3317B67173}' = 'Profile'
            '{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}' = 'PrintHood'
            '{76FC4E2D-D6AD-4519-A663-37BD56068185}' = 'Printers'
            '{DE92C1C7-837F-4F69-A3BB-86E631204A23}' = 'Playlists'
            '{33E28130-4E1E-4676-835A-98395C3BC3BB}' = 'Pictures'
            '{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}' = 'PhotoAlbums'
            '{2C36C0AA-5812-4B87-BFD0-4CD0DFB19B39}' = 'OriginalImages'
            '{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}' = 'Network'
            '{C5ABBF53-E17F-4121-8900-86626FC2C973}' = 'NetHood'
            '{4BD8D571-6D19-48D3-BE97-422220080E43}' = 'Music'
            '{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}' = 'LocalizedResourcesDir'
            '{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}' = 'LocalAppData'
            '{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}' = 'Links'
            '{352481E8-33BE-4251-BA85-6007CAEDCF9D}' = 'InternetCache'
            '{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}' = 'Internet'
            '{D9DC8A3B-B784-432E-A781-5A1130A75963}' = 'History'
            '{054FAE61-4DD8-4787-80B6-090220C4B700}' = 'GameTasks'
            '{CAC52C1A-B53D-4EDC-92D7-6B2E8AC19434}' = 'Games'
            '{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}' = 'Fonts'
            '{1777F761-68AD-4D8A-87BD-30B759FA33DD}' = 'Favorites'
            '{374DE290-123F-4565-9164-39C4925E467B}' = 'Downloads'
            '{FDD39AD0-238F-46AF-ADB4-6C85480369C7}' = 'Documents'
            '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}' = 'Desktop'
            '{2B0F765D-C0E9-4171-908E-08A611B84FF6}' = 'Cookies'
            '{82A74AEB-AEB4-465C-A014-D097EE346D63}' = 'ControlPanel'
            '{56784854-C6CB-462B-8169-88E350ACB882}' = 'Contacts'
            '{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}' = 'Connections'
            '{4BFEFB45-347D-4006-A5BE-AC0CB0567192}' = 'Conflict'
            '{0AC0837C-BBF8-452A-850D-79D08E667CA7}' = 'Computer'
            '{B94237E7-57AC-4347-9151-B08C6C32D1F7}' = 'CommonTemplates'
            '{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}' = 'CommonStartup'
            '{A4115719-D62E-491D-AA7C-E74B8BE3B067}' = 'CommonStartMenu'
            '{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}' = 'CommonPrograms'
            '{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}' = 'CommonOEMLinks'
            '{D0384E7D-BAC3-4797-8F14-CBA229B392B5}' = 'CommonAdminTools'
            '{DF7266AC-9274-4867-8D55-3BD661DE872D}' = 'ChangeRemovePrograms'
            '{9E52AB10-F80D-49DF-ACB8-4330F5687855}' = 'CDBurning'
            '{A305CE99-F527-492B-8B1A-7E76FA98D6E4}' = 'AppUpdates'
            '{A520A1A4-1780-4FF6-BD18-167343C5AF16}' = 'AppDataLow'
            '{724EF170-A42D-4FEF-9F26-B60E846FBA4F}' = 'AdminTools'
            '{DE61D971-5EBC-4F02-A3A9-6C82895E5C04}' = 'AddNewPrograms'
        }
    }

    process {
        foreach ($Session in $CimSession) {
            $ComputerName = $Session.ComputerName
            if (-not $Session.ComputerName) { $ComputerName = 'localhost' }

            # Display a progress activity for each CIM session
            Write-Progress -Id 1 -Activity 'CimSweep - UserAssist sweep' -Status "($($CurrentCIMSession+1)/$($CIMSessionCount)) Current computer: $ComputerName" -PercentComplete (($CurrentCIMSession / $CIMSessionCount) * 100)
            $CurrentCIMSession++

            $CommonArgs = @{}

            if ($Session.Id) { $CommonArgs['CimSession'] = $Session }
            
            $UserSids = Get-HKUSID @CommonArgs
            
            foreach ($Sid in $UserSids) {

                $Parameters = @{
                    Hive = 'HKU'
                    SubKey = "$Sid\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist"
                    Recurse = $true
                }
    
                Get-CSRegistryKey @Parameters @CommonArgs | Where-Object { $_.SubKey -like "*Count" } | Get-CSRegistryValue @CommonArgs | ForEach-Object {
                            
                    # Decrypt Rot13 from https://github.com/StackCrash/PoshCiphers
                    # truncated && streamlined algorithm a little

                    $PlainCharList = New-Object Collections.Generic.List[char]
                    foreach ($CipherChar in $_.ValueName.ToCharArray()) {
    
                        switch ($CipherChar) {
                            { $_ -ge 65 -and $_ -le 90 } { $PlainCharList.Add((((($_ - 65 - 13) % 26 + 26) % 26) + 65)) } # Uppercase characters
                            { $_ -ge 97 -and $_ -le 122 } { $PlainCharList.Add((((($_ - 97 - 13) % 26 + 26) % 26) + 97)) } # Lowercase characters
                            default { $PlainCharList.Add($CipherChar) } # Pass through symbols and numbers
                        }
                    }
                    
                    [string]$Name = -join $PlainCharList

                    # Resolve known folders
                    if ($Name.Length -gt 38) { 
                        $KnownFolderGuid = $Name.Substring(0,38)
                        if ($KnownFolderGuid -match "\{[0-Z]{8}\-[0-Z]{4}\-[0-Z]{4}\-[0-Z]{4}\-[0-Z]{12}\}") { 
                            $KnownFolder = $KnownFolderMapping[$KnownFolderGuid] 
                            if ($KnownFolder) { $Name = $Name.Replace($KnownFolderGuid, $KnownFolder) }
                        }
                    }

                    $ValueContent = $_.ValueContent

                    # Parse LastExecutedTime
                    $FileTime = switch ($ValueContent.Count) {
                              8 { [datetime]::FromFileTime(0) }
                             16 { [datetime]::FromFileTime([BitConverter]::ToInt64($ValueContent[8..15],0)) }
                        default { [datetime]::FromFileTime([BitConverter]::ToInt64($ValueContent[60..67],0)) }
                    }

                    $ObjectProperties = [ordered] @{ 
                        PSTypeName = 'CimSweep.UserAssistEntry'
                        Name = $Name
                        UserSid = $Sid
                        LastExecutedTime = $FileTime.ToUniversalTime().ToString('o')
                    }

                    if ($_.PSComputerName) { $ObjectProperties['PSComputerName'] = $_.PSComputerName }
                    [PSCustomObject]$ObjectProperties
                }
            } 
        }
    }
    end {}
}

Export-ModuleMember -Function Get-CSUserAssist