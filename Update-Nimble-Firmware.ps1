<#
Author: Stan Crider
Date: 11-5-2019
What this crap does:
Updates Nimble to latest specified firmware version
###  Must have HPENimblePowerShellToolkit module installed!!!
###  Nimble uses port 5392 for API calls  ###
###  Must have TXT file of list of Nimble names!!!
###  Must have administrative access to the Nimble interface!!!
#>

#Requires -Module HPENimblePowerShellToolkit

## User Variables
# Location of TXT file of list-of-Nimble-names
$NimbleDeviceList = Get-Content "C:\Temp\Nimble Group List.txt"
# Firmware version of Nimble OS to upgrade to
$NewVer = "5.0.8.0-677726-opt"
# Interface access
$UserName = "admin"

## Script below
# Get password for access
$Credentials = Get-Credential -Message "This script requires administrative access to the Nimble interface." -User $UserName

ForEach($NimbleDevice in $NimbleDeviceList){
    If(Test-Connection $NimbleDevice -Quiet){
        Write-Output "Processing $NimbleDevice..."
        Connect-NSGroup -Group $NimbleDevice -Credential $Credentials -IgnoreServerCertificate

        $Group = Get-NSGroup | Select-Object name,id,version_current
        $VerStatus = Get-NSSoftwareVersion -fields version,status -ErrorAction SilentlyContinue
        $DLV = $VerStatus | Where-Object{$_.name  -eq "downloaded"}
        $Installed = $VerStatus | Where-Object{$_.name  -eq "installed"}

        If($Installed.version -eq $NewVer){
            Write-Output "$NewVer is already installed."
        }
        ElseIf(("downloading" -notin $VerStatus.name) -and ($DLV.version -ne $NewVer)){
            Start-NSGroupSoftwareDownload -id $Group.id -version $NewVer -ErrorAction SilentlyContinue
            Get-NSSoftwareVersion -ErrorAction SilentlyContinue
        }
        ElseIf($DLV.version -eq $NewVer){
            Write-Output "$NewVer is ready to install."
            $RunUpdate = (Start-NSGroupSoftwareUpdate -id $Group.id).array_response_list
            If($RunUpdate.error -eq "SM_ok"){
                Write-Output "$NewVer is being installed."
            }
        }
        Else{
            Write-Output "$NewVer is downloading."
        }

        Disconnect-NSGroup
    }
    Else{
        Write-Warning ("$NimbleDevice is not available.")
    }
    # Output separater
    Write-Output("_"*40)
}
