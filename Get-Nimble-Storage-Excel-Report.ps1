<#
Author: Stan Crider
Date: 3Apr2018
What this crap does:
Create spreadsheet report of Nimble devices from specified list
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
###  Must have HPENimblePowerShellToolkit module installed!!!
###  Nimble uses port 5392 for API calls  ###
#>

#region Functions
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
#endregion

#region Gather Data
# Specify text file with each Nimble group name on separate line
$NimbleDeviceFile = "C:\Temp\Nimble\Nimble Group List.txt"

# Verify device list file
If(Test-Path $NimbleDeviceFile){

# Create Excel output file
    $Date = Get-Date -Format yyyyMMMdd
    $Workbook = "C:\Temp\Nimble\Nimble Reports\Nimble Report $Date.xlsx"

# Create worksheet arrays
    $GroupSheet = @()
    $ArraySheet = @()
    $PoolSheet = @()
    $VolumeSheet = @()
    $DiskSheet = @()
    $RepPartnerSheet = @()
    $NICInterfaceSheet = @()
    $NICConfigSheet = @()
    $NetworkingSheet = @()
    $InitiatorInfoSheet = @()
    $InitGroupSheet = @()
    $ErrorArray = @()

    $NimbleDeviceList = Get-Content $NimbleDeviceFile

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

    ForEach($NimbleDevice in $NimbleDeviceList){
        If(Test-Connection $NimbleDevice -Quiet){
#region Group Initiator
            Write-Output "Processing group $NimbleDevice" # Display on screen to know which array causes error, if any

            Try{
                Connect-NSGroup -Group $NimbleDevice -Credential $Credentials -IgnoreServerCertificate -ErrorAction SilentlyContinue
            }
            Catch{
                $ErrorArray += [PSCustomObject]@{
                    "Group" = $NimbleDevice
                    "Section" = "API Connection"
                    "Error" = $_.Exception.Message
                }
                Continue
            }
            
            $Group = Get-NSGroup -ErrorAction SilentlyContinue
            $Arrays = Get-NSArray -ErrorAction SilentlyContinue
            $Pools = Get-NSPool -ErrorAction SilentlyContinue
            $CHAP = Get-NSChapUser -ErrorAction SilentlyContinue
            $Initiators = Get-NSInitiator -ErrorAction SilentlyContinue
            $Volumes = Get-NSVolume -ErrorAction SilentlyContinue
            $Disks = Get-NSDisk -ErrorAction SilentlyContinue
            $ADMember = Get-NSActiveDirectoryMembership -ErrorAction SilentlyContinue
            $RepPartners = Get-NSReplicationPartner -ErrorAction SilentlyContinue
            $NICConfigs = Get-NSNetworkConfig -ErrorAction SilentlyContinue
            $NICInterfaces = Get-NSNetworkInterface -ErrorAction SilentlyContinue
            $Networks = Get-NSSubnet -ErrorAction SilentlyContinue
            $Initiators = Get-NSInitiator -ErrorAction SilentlyContinue
            $InitGroups = Get-NSInitiatorGroup -ErrorAction SilentlyContinue
            $VerStatus = Get-NSSoftwareVersion -Fields version,status -ErrorAction SilentlyContinue

            $ArrayUsedTotal = 0
            $TotCompressed = 0
            $TotUncompressed = 0
            $TotSnapComp = 0
            $TotSnapUncomp = 0
#endregion

#region Group sheet
                $Sheet1 = "" | Select-Object Group,
                    Domain,
                    Version,
                    AvailableUpdate,
                    DownloadedUpdate,
                    ID,
                    NTP,
                    Timezone,
                    Capacity,
                    FreeSpace,
                    Used,
                    DataCompressed,
                    DataUncompressed,
                    DataRatio,
                    SnapsCompressed,
                    SnapsUncompressed,
                    SnapRatio,
                    TotalCompressed,
                    TotalUncompressed,
                    TotalRatio,
                    RawCapacity,
                    RawDataComp,
                    RawData,
                    RawSnapsComp,
                    RawSnaps,
                    RawTotalComp,
                    RawTotal,
                    PctUsed,
                    Volumes,
                    CHAP,
                    Count,
                    AD,
                    ADName,
                    OU,
                    IQN

                $Sheet1.Group = $Group.name
                $Sheet1.ID = $Group.id
                $Sheet1.Domain = $Group.domain_name
                $Sheet1.NTP = $Group.ntp_server
                $Sheet1.Timezone = $Group.timezone
                $Sheet1.Capacity = Get-Size $Group.usable_capacity_bytes
                $Sheet1.RawCapacity = $Group.usable_capacity_bytes
                $Sheet1.FreeSpace = Get-Size $Group.free_space
                $Sheet1.Volumes = ($Volumes | Measure-Object).Count
                $Sheet1.Version = $Group.version_current
                $Sheet1.IQN = $Group.group_target_name
                $Sheet1.CHAP = $CHAP.name
                $Sheet1.Count = $CHAP.vol_count
                Try{
                    $AvailableUpdates = $VerStatus | Where-Object{$_.name -eq "available"}
                    $DownloadedUpdates = $VerStatus | Where-Object{$_.name -eq "downloaded"}
                }
                Catch{
                    $AvailableUpdates = $null
                    $DownloadedUpdates = $null
                }
                $Sheet1.AvailableUpdate = $AvailableUpdates.Version -join ", "
                $Sheet1.DownloadedUpdate = $DownloadedUpdates.Version -join ", "
                If($ADMember){
                    $Sheet1.AD = "Yes"
                    $Sheet1.ADName = $ADMember.name
                    $Sheet1.OU = $ADMember.organizational_unit
                }
                Else{
                    $Sheet1.AD = "No"
                    $Sheet1.ADName = "N/A"
                    $Sheet1.OU = "N/A"
                }
#end region

# Array sheet
                ForEach($Array in $Arrays){
                    $ArrayUsedTotal += $Array.usage

                    $Sheet2 = "" | Select-Object Group,
                        Array,
                        Model,
                        FullModel,
                        Version,
                        Serial,
                        Pool,
                        Role,
                        GB_NICs,
                        TenGB_NICs,
                        Capacity,
                        FreeSpace,
                        DedupCapacity,
                        DedupUsage

                    $Sheet2.Group = $Group.name
                    $Sheet2.Array = $Array.full_name
                    $Sheet2.Model = $Array.model
                    $Sheet2.FullModel = $Array.extended_model
                    $Sheet2.Version = $Array.version
                    $Sheet2.Serial = $Array.serial
                    $Sheet2.Pool = $Array.pool_name
                    $Sheet2.Role = $Array.role
                    $Sheet2.GB_NICs = $Array.gig_nic_port_count
                    $Sheet2.TenGB_NICs = $Array.ten_gig_sfp_nic_port_count
                    $Sheet2.Capacity = Get-Size $Array.raw_capacity_bytes
                    $Sheet2.FreeSpace = Get-Size $Array.available_bytes
                    $Sheet2.DedupCapacity = Get-Size $Array.dedupe_capacity_bytes
                    $Sheet2.DedupUsage = Get-Size $Array.dedupe_usage_bytes

                    $ArraySheet += $Sheet2

                }
                $Sheet1.Used = Get-Size $ArrayUsedTotal
                $GroupSheet += $Sheet1

    # Pool sheet -- Pool sheet only gets created if more than one pool exists in a group
                If((($Pools | Measure-Object).Count) -gt 1){
                    ForEach($Pool in $Pools){
                        $Sheet3 = "" | Select-Object Group,
                        Pool,
                        Capacity,
                        Used,
                        Free,
                        Default,
                        VolCount,
                        SnapCount

                        $Sheet3.Group = $Group.name
                        $Sheet3.Pool = $Pool.name
                        $Sheet3.Capacity = Get-Size $Pool.capacity
                        $Sheet3.Used = Get-Size $Pool.usage
                        $Sheet3.Free = Get-Size $Pool.free_space
                        $Sheet3.Default = $Pool.is_default
                        $Sheet3.VolCount = $Pool.vol_count
                        $Sheet3.SnapCount = $Pool.snap_count
                        $PoolSheet += $Sheet3
                    }
                }

    # Volume sheet
                ForEach($Volume in $Volumes){
                    $Sheet4 = "" | Select-Object Group,
                        Owner,
                        Volume,
                        Serial,
                        Online,
                        MultiInit,
                        Dedupe,
                        Size,
                        Used,
                        Compressed,
                        Uncompressed,
                        SnapshotCompressed,
                        SnapshotUncompressed

                    $Sheet4.Group = $Group.name
                    $Sheet4.Owner = $Volume.owned_by_group
                    $Sheet4.Volume = $Volume.name
                    $Sheet4.Serial = $Volume.serial_number
                    $Sheet4.Online = $Volume.vol_state
                    $Sheet4.MultiInit = $Volume.multi_initiator
                    $Sheet4.Dedupe = $Volume.dedupe_enabled
                    $Sheet4.Size = Get-Size ($Volume.size*1MB) #size attribute is in MB, not raw bytes; multiply by MB to run through converter function
                    $Sheet4.Used = Get-Size $Volume.total_usage_bytes
                    $Sheet4.Compressed = Get-Size $Volume.vol_usage_compressed_bytes
                    $Sheet4.Uncompressed = Get-Size $Volume.vol_usage_uncompressed_bytes
                    $Sheet4.SnapshotCompressed = Get-Size $Volume.snap_usage_compressed_bytes
                    $Sheet4.SnapshotUncompressed = Get-Size $Volume.snap_usage_uncompressed_bytes
                    $VolumeSheet += $Sheet4
                    $TotCompressed += $Volume.vol_usage_compressed_bytes
                    $TotUncompressed += $Volume.vol_usage_uncompressed_bytes
                    $TotSnapComp += $Volume.snap_usage_compressed_bytes
                    $TotSnapUncomp += $Volume.snap_usage_uncompressed_bytes
                }

    # Divide by zero error handling for group data sizes
                If($TotCompressed -eq 0){
                    $DataRatio = 0
                }
                Else{
                    $DataRatio = [math]::Round(($TotUncompressed/$TotCompressed),2)
                }

                If($TotSnapComp -eq 0){
                    $SnapRatio = 0
                }
                Else{
                    $SnapRatio = [math]::Round(($TotSnapUncomp/$TotSnapComp),2)
                }

                If(($TotCompressed + $TotSnapComp) -eq 0){
                    $TotalRatio = 0
                }
                Else{
                    $TotalRatio = [math]::Round(($TotUncompressed + $TotSnapUncomp)/($TotCompressed + $TotSnapComp),2)
                }

    # Group sheet data sizes
                $Sheet1.DataCompressed = Get-Size $TotCompressed
                $Sheet1.DataUncompressed = Get-Size $TotUncompressed
                $Sheet1.DataRatio = $DataRatio
                $Sheet1.SnapsCompressed = Get-Size $TotSnapComp
                $Sheet1.SnapsUncompressed = Get-Size $TotSnapUncomp
                $Sheet1.SnapRatio = $SnapRatio
                $Sheet1.TotalCompressed = Get-Size ($TotCompressed + $TotSnapComp)
                $Sheet1.TotalUncompressed = Get-Size ($TotUncompressed + $TotSnapUncomp)
                $Sheet1.TotalRatio = $TotalRatio
                $Sheet1.RawData = $TotUncompressed
                $Sheet1.RawDataComp = $TotCompressed
                $Sheet1.RawSnaps = $TotSnapUncomp
                $Sheet1.RawSnapsComp = $TotSnapComp
                $Sheet1.RawTotal = ($TotUncompressed + $TotSnapUncomp)
                $Sheet1.RawTotalComp = ($TotCompressed + $TotSnapComp)
                $Sheet1.PctUsed = [math]::Round((($TotUncompressed + $TotSnapUncomp)/$Group.usable_capacity_bytes),2)

    # Disk sheet
                ForEach($Disk in $Disks){
                    $Sheet5 = "" | Select-Object Group,
                        Array,
                        ShelfID,
                        Location,
                        Serial,
                        Slot,
                        Bank,
                        State,
                        Size,
                        Type,
                        Model,
                        Vendor,
                        Firmware

                    $Sheet5.Group = $Group.name
                    $Sheet5.Array = $Disk.array_name
                    $Sheet5.ShelfID = $Disk.shelf_id
                    $Sheet5.Location = $Disk.shelf_location
                    $Sheet5.Serial = $Disk.shelf_serial
                    $Sheet5.Slot = $Disk.slot
                    $Sheet5.Bank = $Disk.bank
                    $Sheet5.State = $Disk.state
                    $Sheet5.Size = Get-Size $Disk.size
                    $Sheet5.Type = $Disk.type
                    $Sheet5.Model = $Disk.model
                    $Sheet5.Vendor = $Disk.vendor
                    $Sheet5.Firmware = $Disk.firmware_version
                    $DiskSheet += $Sheet5
                }

    # Replication sheet
                ForEach($RepPartner in $RepPartners){
                    $Sheet6 = "" | Select-Object Group,
                        SyncStatus,
                        PartnerName,
                        HostName,
                        Description,
                        ArraySerial,
                        Alive,
                        EnableMatch,
                        PartnerType,
                        Paused,
                        PoolName,
                        RepDirection,
                        Version,
                        VolColListCount

                    $Sheet6.Group = $Group.name
                    $Sheet6.SyncStatus = $RepPartner.cfg_sync_status
                    $Sheet6.PartnerName = $RepPartner.full_name
                    $Sheet6.HostName = $RepPartner.hostname
                    $Sheet6.Description = $RepPartner.description
                    $Sheet6.ArraySerial = $RepPartner.array_serial
                    $Sheet6.Alive = $RepPartner.is_alive
                    $Sheet6.EnableMatch = $RepPartner.match_folder
                    $Sheet6.PartnerType = $RepPartner.partner_type
                    $Sheet6.Paused = $RepPartner.paused
                    $Sheet6.PoolName = $RepPartner.pool_name
                    $Sheet6.RepDirection = $RepPartner.replication_direction
                    $Sheet6.Version = $RepPartner.version
                    $Sheet6.VolColListCount = $RepPartner.volume_collection_list_count

                    $RepPartnerSheet += $Sheet6
                }

    # NIC Config sheet
                ForEach($NICConfig in $NICConfigs){
                    $Sheet7 = "" | Select-Object Group,
                        Manager,
                        MgmtIP,
                        NIC,
                        Role

                    $Sheet7.Group = $Group.name
                    $Sheet7.Manager = $NICConfig.group_leader_array
                    $Sheet7.MgmtIP = $NICConfig.mgmt_ip
                    $Sheet7.NIC = $NICConfig.name
                    $Sheet7.Role = $NICConfig.role

                    $NICConfigSheet += $Sheet7
                }

    # NIC Interface sheet
                ForEach($NICInterface in $NICInterfaces){
                    $Sheet8 = "" | Select-Object Group,
                        Array,
                        Controller,
                        NIC,
                        MAC,
                        Status,
                        Speed

                    $Sheet8.Group = $Group.name
                    $Sheet8.Array = $NICInterface.array_name_or_serial
                    $Sheet8.Controller = $NICInterface.controller_name
                    $Sheet8.NIC = $NICInterface.name
                    $Sheet8.MAC = $NICInterface.mac
                    $Sheet8.Status = $NICInterface.link_status
                    $Sheet8.Speed = $NICInterface.link_speed

                    $NICInterfaceSheet += $Sheet8
                }

    # Network sheet
                ForEach($Network in $Networks){
                    $Sheet9 = "" | Select-Object Group,
                        IP,
                        Name,
                        Mask,
                        Network,
                        MTU,
                        GroupAllow,
                        iSCSI,
                        Type,
                        VLAN,
                        NetZone

                    $Sheet9.Group = $Group.name
                    $Sheet9.IP = $Network.discovery_ip
                    $Sheet9.Name = $Network.name
                    $Sheet9.Mask = $Network.netmask
                    $Sheet9.Network = $Network.network
                    $Sheet9.MTU = $Network.mtu
                    $Sheet9.GroupAllow = $Network.allow_group
                    $Sheet9.iSCSI = $Network.allow_iscsi
                    $Sheet9.Type = $Network.type
                    $Sheet9.VLAN = $Network.vlan_id
                    $Sheet9.NetZone = $Network.netzone_type

                    $NetworkingSheet += $Sheet9
                }

    # Initiator sheet
                ForEach($Initiator in $Initiators){
                    $Sheet10 = "" | Select-Object Group,
                        InitGroup,
                        Protocol,
                        Label,
                        IP,
                        IQN

                    $Sheet10.Group = $Group.name
                    $Sheet10.InitGroup = $Initiator.initiator_group_name
                    $Sheet10.Protocol = $Initiator.access_protocol
                    $Sheet10.Label = $Initiator.label
                    $Sheet10.IP = $Initiator.ip_address
                    $Sheet10.IQN = $Initiator.iqn

                    $InitiatorInfoSheet +=$Sheet10
                }

    #  Initiator group sheet
                ForEach($InitGroup in $InitGroups){
                    $Sheet11 = "" | Select-Object Group,
                        InitGroup,
                        Connections,
                        Volumes,
                        Protocol,
                        HostType

                    $Sheet11.Group = $Group.name
                    $Sheet11.InitGroup = $InitGroup.name
                    $Sheet11.Connections = $InitGroup.num_connections
                    $Sheet11.Volumes = $InitGroup.volume_count
                    $Sheet11.Protocol = $InitGroup.access_protocol
                    $Sheet11.HostType = $InitGroup.host_type

                    $InitGroupSheet += $Sheet11
                }

                Disconnect-NSGroup

        }
        Else{
            Write-Warning "Group $NimbleDevice is not responding."
            $ErrorArray += [PSCustomObject]@{
                "Group" = $NimbleDevice
                "Section" = "Network Connection"
                "Error" = "Group is not responding."
            }
        }
    }
#endregion

#region Output to Excel file
    # Create Excel standard configuration properties
    $ExcelProps = @{
        Autosize = $true;
        FreezeTopRow = $true;
        BoldTopRow = $true;
    }

    $ExcelProps.Path = $Workbook

    # Group sheet
    $GroupSheetLastRow = ($GroupSheet | Measure-Object).Count + 1
    If($GroupSheetLastRow -gt 1){
        $GroupSheetHeaderCount = Get-ColumnName ($GroupSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $GroupSheetHeaderRow = "'Groups'!`$A`$1:`$$GroupSheetHeaderCount`$1"
        $GroupSheetNumberColumns = "'Groups'!`$U`$2:`$AA`$$GroupSheetLastRow"
        $GroupSheetStyle = @()
        $GroupSheetStyle += New-ExcelStyle -Range $GroupSheetHeaderRow -HorizontalAlignment Center
        $GroupSheetStyle += New-ExcelStyle -Range $GroupSheetNumberColumns -NumberFormat '0'
        $GroupSheet | Export-Excel @ExcelProps -WorkSheetname "Groups" -Style $GroupSheetStyle
    }
    # Array sheet
    $ArraySheetLastRow = ($ArraySheet | Measure-Object).Count + 1
    If($ArraySheetLastRow -gt 1){
        $ArraySheetHeaderCount = Get-ColumnName ($ArraySheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $ArraySheetHeaderRow = "'Arrays'!`$A`$1:`$$ArraySheetHeaderCount`$1"
        $ArraySheetStyle = New-ExcelStyle -Range $ArraySheetHeaderRow -HorizontalAlignment Center
        $ArraySheet | Export-Excel @ExcelProps -WorkSheetname "Arrays" -Style $ArraySheetStyle
    }

    # Pool sheet
    $PoolSheetLastRow = ($PoolSheet | Measure-Object).Count + 1
    If($PoolSheetLastRow -gt 1){
        $PoolSheetHeaderCount = Get-ColumnName ($PoolSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $PoolSheetHeaderRow = "'Pools'!`$A`$1:`$$PoolSheetHeaderCount`$1"
        $PoolSheetStyle = New-ExcelStyle -Range $PoolSheetHeaderRow -HorizontalAlignment Center
        $PoolSheet | Export-Excel @ExcelProps -WorkSheetname "Pools" -Style $PoolSheetStyle
    }

    # Volume sheet
    $VolumeSheetLastRow = ($VolumeSheet | Measure-Object).Count + 1
    If($VolumeSheetLastRow -gt 1){
        $VolumeSheetHeaderCount = Get-ColumnName ($VolumeSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $VolumeSheetHeaderRow = "'Volumes'!`$A`$1:`$$VolumeSheetHeaderCount`$1"
        $VolumeSheetStyle = New-ExcelStyle -Range $VolumeSheetHeaderRow -HorizontalAlignment Center
        $VolumeSheet | Export-Excel @ExcelProps -WorkSheetname "Volumes" -Style $VolumeSheetStyle
    }

    # Disk sheet
    $DiskSheetLastRow = ($DiskSheet | Measure-Object).Count + 1
    If($DiskSheetLastRow -gt 1){
        $DiskSheetHeaderCount = Get-ColumnName ($DiskSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $DiskSheetHeaderRow = "'Disks'!`$A`$1:`$$DiskSheetHeaderCount`$1"
        $DiskSheetStyle = New-ExcelStyle -Range $DiskSheetHeaderRow -HorizontalAlignment Center
        $DiskSheet | Export-Excel @ExcelProps -WorkSheetname "Disks" -Style $DiskSheetStyle
    }

    # Replication partner sheet
    $RepPartnerSheetLastRow = ($RepPartnerSheet | Measure-Object).Count + 1
    If($RepPartnerSheetLastRow -gt 1){
        $RepPartnerSheetHeaderCount = Get-ColumnName ($RepPartnerSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $RepPartnerSheetHeaderRow = "'RepPartners'!`$A`$1:`$$RepPartnerSheetHeaderCount`$1"
        $RepPartnerSheetStyle = New-ExcelStyle -Range $RepPartnerSheetHeaderRow -HorizontalAlignment Center
        $RepPartnerSheet | Export-Excel @ExcelProps -WorkSheetname "RepPartners" -Style $RepPartnerSheetStyle
    }

    # NIC config sheet
    $NICConfigSheetLastRow = ($NICConfigSheet | Measure-Object).Count + 1
    If($NICConfigSheetLastRow -gt 1){
        $NICConfigSheetHeaderCount = Get-ColumnName ($NICConfigSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $NICConfigSheetHeaderRow = "'NIC Configs'!`$A`$1:`$$NICConfigSheetHeaderCount`$1"
        $NICConfigSheetStyle = New-ExcelStyle -Range $NICConfigSheetHeaderRow -HorizontalAlignment Center
        $NICConfigSheet | Export-Excel @ExcelProps -WorkSheetname "NIC Config" -Style $NICConfigSheetStyle
    }

    # NIC interface sheet
    $NICInterfaceSheetLastRow = ($NICInterfaceSheet | Measure-Object).Count + 1
    If($NICInterfaceSheetLastRow -gt 1){
        $NICInterfaceSheetHeaderCount = Get-ColumnName ($NICInterfaceSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $NICInterfaceSheetHeaderRow = "'NIC Interface'!`$A`$1:`$$NICInterfaceSheetHeaderCount`$1"
        $NICInterfaceSheetStyle = New-ExcelStyle -Range $NICInterfaceSheetHeaderRow -HorizontalAlignment Center
        $NICInterfaceSheet | Export-Excel @ExcelProps -WorkSheetname "NIC Interface" -Style $NICInterfaceSheetStyle
    }

    # Networking sheet
    $NetworkingSheetLastRow = ($NetworkingSheet | Measure-Object).Count + 1
    If($NetworkingSheetLastRow -gt 1){
        $NetworkingSheetHeaderCount = Get-ColumnName ($NetworkingSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $NetworkingSheetHeaderRow = "'Networking'!`$A`$1:`$$NetworkingSheetHeaderCount`$1"
        $NetworkingSheetStyle = New-ExcelStyle -Range $NetworkingSheetHeaderRow -HorizontalAlignment Center
        $NetworkingSheet | Export-Excel @ExcelProps -WorkSheetname "Networking" -Style $NetworkingSheetStyle
    }

    # Initiator sheet
    $InitiatorInfoSheetLastRow = ($InitiatorInfoSheet | Measure-Object).Count + 1
    If($InitiatorInfoSheetLastRow -gt 1){
        $InitiatorSheetHeaderCount = Get-ColumnName ($InitiatorInfoSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $InitiatorSheetHeaderRow = "'Initiator'!`$A`$1:`$$InitiatorSheetHeaderCount`$1"
        $InitiatorInfoSheetStyle = New-ExcelStyle -Range $InitiatorSheetHeaderRow -HorizontalAlignment Center
        $InitiatorInfoSheet | Export-Excel @ExcelProps -WorkSheetname "Initiator" -Style $InitiatorInfoSheetStyle
    }

    # Initiator group sheet
    $InitGroupSheetLastRow = ($InitGroupSheet | Measure-Object).Count + 1
    If($InitGroupSheetLastRow -gt 1){
        $InitGroupSheetHeaderCount = Get-ColumnName ($InitGroupSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $InitGroupSheetHeaderRow = "'InitGroup'!`$A`$1:`$$InitGroupSheetHeaderCount`$1"
        $InitGroupSheetStyle = New-ExcelStyle -Range $InitGroupSheetHeaderRow -HorizontalAlignment Center
        $InitGroupSheet | Export-Excel @ExcelProps -WorkSheetname "InitGroup" -Style $InitGroupSheetStyle
    }

    # Error sheet
    If($ErrorArray -ne ""){
        $ErrorArrayHeaderCount = Get-ColumnName ($ErrorArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $ErrorArrayHeaderRow = "`$A`$1:`$$ErrorArrayHeaderCount`$1"
        $ErrorArrayStyle = New-ExcelStyle -Range "Errors$ErrorArrayHeaderRow" -HorizontalAlignment Center
        $ErrorArray | Export-Excel @ExcelProps -WorkSheetname "Errors" -Style $ErrorArrayStyle
    }
}
#endregion

# Error handling for source file location
Else{
    Write-Warning "The file $NimbleDeviceFile is not valid. Check the file name and try again."
}
