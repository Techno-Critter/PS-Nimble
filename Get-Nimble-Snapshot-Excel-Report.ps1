<#
Author: Stan Crider
Date: 30May2018
What this crap does:
Get Nimble Snapshot report from specified Nimble Group list
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
###  Must have HPENimblePowerShellToolkit module installed!  ###
###  Nimble uses port 5392 for API calls  ###
#>

#Requires -Module HPENimblePowerShellToolkit
#Requires -Module ImportExcel

# Function: Change data sizes to legible values; converts number to string
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

# Specify text file with each Nimble group name on separate line
$NimbleDeviceFile = "C:\Temp\Nimble\Nimble Group List.txt"

# Verify device list file
If(Test-Path $NimbleDeviceFile){

# Create Excel output file
    $Date = Get-Date -Format yyyyMMMdd
    $Workbook = ("C:\Temp\Nimble\Nimble Reports\Nimble Replication Report $Date.xlsx")

# Set arrays
    $ErrorArray = @()

# Set username and password to access each group; NOTE: account must have permissions to each group!
    $Credentials = Get-Credential -UserName "admin" -Message "Nimble Controller Credentials:"

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
            $SnapShots | Export-Excel -Path $Workbook -WorkSheetname $WSName.name -BoldTopRow -AutoSize -FreezeTopRow

            Disconnect-NSGroup
        }
        Else{
            Write-Warning "Device $NimbleDevice is not responding."
            $ErrorArray += [PSCustomObject]@{
                "Device" = $NimbleDevice
                "Error" = "Device not responding"
            }
        }
    }
    $ErrorArray | Export-Excel -Path $Workbook -WorksheetName "Errors" -BoldTopRow -AutoSize -FreezeTopRow

}

# Error handling for source file location
Else{
    Write-Warning "The file $NimbleDeviceFile is not valid. Check the file name and try again."
}
