<#
Author: Stan Crider
Date: 30May2018
What this crap does:
Get Nimble Snapshot report from specified Nimble Group list
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
###  Must have HPENimblePowerShellToolkit module installed!!!
###  Nimble uses port 5392 for API calls  ###
#>

#Requires -Module HPENimblePowerShellToolkit
#Requires -Module ImportExcel

## FUNCTIONS
# Change data sizes to legible values; converts number to string
Function Get-Size([double]$DataSize){
    Switch($DataSize){
        {$_ -lt 1KB}{
            $DataValue =  "$DataSize B"
        }
        {($_ -ge 1KB) -and ($_ -lt 1MB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1KB) + " KiB"
        }
        {($_ -ge 1MB) -and ($_ -lt 1GB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1MB) + " MiB"
        }
        {($_ -ge 1GB) -and ($_ -lt 1TB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1GB) + " GiB"
        }
        {($_ -ge 1TB) -and ($_ -lt 1PB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1TB) + " TiB"
        }
        Default{
            $DataValue = "{0:N2}" -f ($DataSize/1PB) + " PiB"
        }
    }
    $DataValue
}

# Convert number of object items into Excel column headers
Function Get-ColumnName ([int]$ColumnCount){
    If(($ColumnCount -le 702) -and ($ColumnCount -ge 1)){
        $ColumnCount = [Math]::Floor($ColumnCount)
        $CharStart = 64
        $FirstCharacter = $null

        # Convert number into double letter column name (AA-ZZ)
        If($ColumnCount -gt 26){
            $FirstNumber = [Math]::Floor(($ColumnCount)/26)
            $SecondNumber = ($ColumnCount) % 26

            # Reset increment for base-26
            If($SecondNumber -eq 0){
                $FirstNumber--
                $SecondNumber = 26
            }

            # Left-side column letter (first character from left to right)
            $FirstLetter = [int]($FirstNumber + $CharStart)
            $FirstCharacter = [char]$FirstLetter

            # Right-side column letter (second character from left to right)
            $SecondLetter = $SecondNumber + $CharStart
            $SecondCharacter = [char]$SecondLetter

            # Combine both letters into column name
            $CharacterOutput = $FirstCharacter + $SecondCharacter
        }

        # Convert number into single letter column name (A-Z)
        Else{
            $CharacterOutput = [char]($ColumnCount + $CharStart)
        }
    }
    Else{
        $CharacterOutput = "ZZ"
    }

    # Output column name
    $CharacterOutput
}

# Specify text file with each Nimble group name on separate line
$NimbleDeviceFile = "C:\Temp\Nimble\Nimble Group List.txt"

# Verify device list file
If(Test-Path $NimbleDeviceFile){

# Create Excel output file
    $Date = Get-Date -Format yyyyMMMdd
    $Workbook = ("C:\Temp\Nimble\Nimble Reports\Nimble Replication Report $Date.xlsx")

# Create Excel standard configuration properties
    $ExcelProps = @{
        Autosize = $true;
        FreezeTopRow = $true;
        BoldTopRow = $true;
    }

    $ExcelProps.Path = $Workbook

# Set arrays
    $ErrorArray = @()

# Set username and password to access each group; NOTE: account must have permissions to each group!
    $NimbleAccount = "admin"
    $Hostname = $ENV:COMPUTERNAME
    $CurrentUser = $ENV:USERNAME
    $CredentialFileDirectory = "C:\Credential Files"
    $CredentialFile = "$CredentialFileDirectory\$Hostname\NimbleCreds $CurrentUser.xml"
    If(Test-Path $CredentialFile){
        $Credentials = Import-Clixml $CredentialFile
    }
    Else{
        $Credentials = Get-Credential -UserName $NimbleAccount -Message "Provide the password for the Nimble account: $NimbleAccount"
        If(-Not (Test-Path "$CredentialFileDirectory\$Hostname")){
            New-Item -Path $CredentialFileDirectory -Name $Hostname -ItemType Directory
        }
        $Credentials | Export-Clixml $CredentialFile
    }

    $NimbleDeviceList = Get-Content $NimbleDeviceFile

    ForEach($NimbleDevice in $NimbleDeviceList){
        If(Test-Connection $NimbleDevice -Quiet){
            Write-Output("Processing group " + $NimbleDevice) # Display on screen to know which array causes error, if any
            Try{
                Connect-NSGroup -Group $NimbleDevice -Credential $Credentials -IgnoreServerCertificate
            }
            Catch{
                $ErrorArray += [PSCustomObject]@{
                    "Device" = $NimbleDevice
                    "Error" = $_.Exception.Message
                }
                Continue
            }
            $WSName = Get-NSGroup | Select-Object name

            $SnapShots = Get-NSSnapshotCollection | Select-Object @{Name="Snapshot";Expression={$_.name}},
                    @{Name="Manual";Expression={$_.is_manual}},
                    @{Name="Schedule";Expression={$_.sched_name}},
                    @{Name="VolCollection";Expression={$_.volcoll_name}},
                    @{Name="Complete";Expression={$_.is_complete}},
                    @{Name="Origin";Expression={$_.origin_name}},
                    @{Name="Unmanaged";Expression={$_.is_unmanaged}},
                    @{Name="Status";Expression={$_.online_status}},
                    @{Name="ReplEnabled";Expression={$_.replicate}},
                    @{Name="ReplStatus";Expression={$_.repl_status}},
                    @{Name="DataTransferred";Expression={Get-Size $_.repl_bytes_transferred}}

            $SnapShotsHeaderCount = Get-ColumnName ($SnapShots | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
            $SnapShotsHeaderRow = "`$A`$1:`$$SnapShotsHeaderCount`$1"
            $SnapShotsStyle = New-ExcelStyle -Range "'Permissions'$SnapShotsHeaderRow" -HorizontalAlignment Center
            $SnapShots | Export-Excel @ExcelProps -WorkSheetname $WSName.name -Style $SnapShotsStyle

            Disconnect-NSGroup
        }
        Else{
            Write-Warning "Device $NimbleDevice is not responding."
            $ErrorArray += [PSCustomObject]@{
                "Device" = $NimbleDevice
                "Error"  = "Device not responding"
            }
        }
    }
    $ErrorArrayHeaderCount = Get-ColumnName ($ErrorArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ErrorArrayHeaderRow = "`$A`$1:`$$ErrorArrayHeaderCount`$1"
    $ErrorArrayStyle = New-ExcelStyle -Range "'Permissions'$ErrorArrayHeaderRow" -HorizontalAlignment Center
    $ErrorArray | Export-Excel @ExcelProps -WorksheetName "Errors" -Style $ErrorArrayStyle

}

# Error handling for source file location
Else{
    Write-Warning "The file $NimbleDeviceFile is not valid. Check the file name and try again."
}
