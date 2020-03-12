<#
Author: Stan Crider
Date: 3Apr2018
What this crap does:
Create spreadsheet report of Nimble devices from specified list
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
###  Must have HPENimblePowerShellToolkit module installed!  ###
###  Nimble uses port 5392 for API calls  ###
#>

#region Function: Change data sizes to legible values; converts number to string
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
    $Credentials = Get-Credential -UserName "admin" -Message "Nimble Controller Credentials:"

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
            $Group = Get-NSGroup
            $Arrays = Get-NSArray
            $Pools = Get-NSPool
            $CHAP = Get-NSChapUser
            $Initiators = Get-NSInitiator
            $Volumes = Get-NSVolume
            $Disks = Get-NSDisk
            $ADMember = Get-NSActiveDirectoryMembership
            $RepPartners = Get-NSReplicationPartner
            $NICConfigs = Get-NSNetworkConfig
            $NICInterfaces = Get-NSNetworkInterface
            $Networks = Get-NSSubnet
            $Initiators = Get-NSInitiator
            $InitGroups = Get-NSInitiatorGroup
            $VerStatus = Get-NSSoftwareVersion -Fields version,status -ErrorAction SilentlyContinue
            <#Try{
                $VerStatus = Get-NSSoftwareVersion -Fields version,status -ErrorAction SilentlyContinue
            }
            Catch{
                $ErrorArray += [PSCustomObject]@{
                    "Group" = $NimbleDevice
                    "Section" = "Version Status"
                    "Error" = $_.Exception.Message
                }
                $VerStatus = $null
            }#>
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
    $HeaderRow = ("!`$A`$1:`$AZ`$1")

# Group sheet
    $GroupSheetLastRow = ($GroupSheet | Measure-Object).Count + 1
    If($GroupSheetLastRow -gt 1){
        $GroupSheetNumberColumns = "Groups!`$U`$2:`$AA`$$GroupSheetLastRow"
        $GroupSheetStyle = @()
        $GroupSheetStyle += New-ExcelStyle -Range "Groups$HeaderRow" -HorizontalAlignment Center
        $GroupSheetStyle += New-ExcelStyle -Range $GroupSheetNumberColumns -NumberFormat '0'
        $GroupSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -Autosize -WorkSheetname "Groups" -Style $GroupSheetStyle
    }
# Array sheet
    $ArraySheetStyle = New-ExcelStyle -Range "Arrays$HeaderRow" -HorizontalAlignment Center
    $ArraySheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -Autosize -WorkSheetname "Arrays" -Style $ArraySheetStyle

# Pool sheet
    $PoolSheetStyle = New-ExcelStyle -Range "Pools$HeaderRow" -HorizontalAlignment Center
    $PoolSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -Autosize -WorkSheetname "Pools" -Style $PoolSheetStyle

# Volume sheet
    $VolumeSheetStyle = New-ExcelStyle -Range "Volumes$HeaderRow" -HorizontalAlignment Center
    $VolumeSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -Autosize -WorkSheetname "Volumes" -Style $VolumeSheetStyle

# Disk sheet
    $DiskSheetStyle = New-ExcelStyle -Range "Disks$HeaderRow" -HorizontalAlignment Center
    $DiskSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -Autosize -WorkSheetname "Disks" -Style $DiskSheetStyle

# Replication partner sheet
    $RepPartnerSheetStyle = New-ExcelStyle -Range "RepPartners$HeaderRow" -HorizontalAlignment Center
    $RepPartnerSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "RepPartners" -Style $RepPartnerSheetStyle

# NIC config sheet
    $NICConfigSheetStyle = New-ExcelStyle -Range "NIC Configs$HeaderRow" -HorizontalAlignment Center
    $NICConfigSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "NIC Config" -Style $NICConfigSheetStyle

# NIC interface sheet
    $NICInterfaceSheetStyle = New-ExcelStyle -Range "NIC Interface$HeaderRow" -HorizontalAlignment Center
    $NICInterfaceSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "NIC Interface" -Style $NICInterfaceSheetStyle

# Networking sheet
    $NetworkingsheetStyle = New-ExcelStyle -Range "Networking$HeaderRow" -HorizontalAlignment Center
    $NetworkingSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "Networking" -Style $NetworkingsheetStyle

# Initiator sheet
    $InitiatorInfoSheetStyle = New-ExcelStyle -Range "Initiator$HeaderRow" -HorizontalAlignment Center
    $InitiatorInfoSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "Initiator" -Style $InitiatorInfoSheetStyle

# Initiator group sheet
    $InitGroupSheetStyle = New-ExcelStyle -Range "InitGroup$HeaderRow" -HorizontalAlignment Center
    $InitGroupSheet | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "InitGroup" -Style $InitGroupSheetStyle

# Error sheet
    If($ErrorArray -ne ""){
        $ErrorArrayStyle = New-ExcelStyle -Range "Errors$HeaderRow" -HorizontalAlignment Center
        $ErrorArray | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorkSheetname "Errors" -Style $ErrorArrayStyle
    }
}
#endregion

# Error handling for source file location
Else{
    Write-Warning "The file $NimbleDeviceFile is not valid. Check the file name and try again."
}
