function Get-TervisClusterNodeToHostVM {
    param(
        [Parameter(Mandatory)]$VMSize,
        [Parameter(Mandatory)]$Cluster
    )
    $ClusterNodes = Get-TervisClusterNode -Cluster $Cluster

    Test-TervisClusterNodesNPlusOne -ClusterNodes $ClusterNodes

    $ClusterNodeWithBestVCPUtoCPURatioAndEnoughFreeMemory = $ClusterNodes |
    where { ($_.FreeMemory - $VMSize.Memory) -gt 32 } |
    sort VirtualCoresToPhysicalCoresRatio |
    select -First 1    

    $ClusterNodeWithBestVCPUtoCPURatioAndEnoughFreeMemory
}

function Test-TervisClusterNodesNPlusOne {
    param(
        [Parameter(Mandatory)]$ClusterNodes
    )
    $ClusterNodeWithTheMostMemory = $ClusterNodes | 
    sort TotalMemory -Descending | 
    select -First 1

    $TotalMemoryOfAllClusterNodes = $ClusterNodes | measure -Sum TotalMemory | select -ExpandProperty Sum
    $TotalUsedMemoryOfAllClusterNodes = $ClusterNodes | measure -Sum UsedMemory | select -ExpandProperty Sum
    if ($TotalUsedMemoryOfAllClusterNodes -gt ($TotalMemoryOfAllClusterNodes - $ClusterNodeWithTheMostMemory.TotalMemory)) {Throw "Cannot create VM as we do not have enough memory in the cluster to be N+1"}
}

function Get-TervisClusterNode {
     param(
        [Parameter(Mandatory)][String]$Cluster
    )
    process {
        $ClusterNodes = Get-ClusterNode -cluster $Cluster
        $ClusterNodes | Mixin-ClusterNode
        $ClusterNodes
    }   
}

filter Mixin-ClusterNode {
    $_ | Add-Member -MemberType ScriptProperty -Name FreeMemory -Value { (Get-WmiObject win32_operatingsystem -ComputerName $This).FreePhysicalMemory / 1MB -as [int] }
    $_ | Add-Member -MemberType ScriptProperty -Name TotalMemory -Value { (Get-WmiObject win32_operatingsystem -ComputerName $This).TotalVisibleMemorySize / 1MB -as [int] }
    $_ | Add-Member -MemberType ScriptProperty -Name UsedMemory -Value { $This.TotalMemory - $this.FreeMemory }
    $_ | Add-Member -MemberType ScriptProperty -Name MemoryPercentFree -Value { ($this.FreeMemory / $This.TotalMemory) * 100 }
    $_ | Add-Member -MemberType ScriptProperty -Name TotalCores -Value { (Get-WmiObject -ComputerName $This Win32_Processor) | Measure-Object -Property NumberOfCores -Sum | select -ExpandProperty sum }
    $_ | Add-Member -MemberType ScriptProperty -Name TotalVirtualCores -Value { (Get-WmiObject -ComputerName $This CIM_Processor -Namespace root\virtualization\v2).count }
    $_ | Add-Member -MemberType ScriptProperty -Name VirtualCoresToPhysicalCoresRatio -Value { $This.TotalVirtualCores / $This.TotalCores }
    $_ | Add-Member -MemberType ScriptProperty -Name ADSite -Value { Get-ComputerSite -ComputerName $this.Name }
}

#http://www.powershellmagazine.com/2013/04/23/pstip-get-the-ad-site-name-of-a-computer/
function Get-ComputerSite {
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][String]$ComputerName
    )
    process {
        $site = nltest /server:$ComputerName /dsgetsite 2>$null
        if($LASTEXITCODE -eq 0){ $site[0] }
    }
}

function Get-ClusterADSite {
    param(
        [Parameter(Mandatory)][string]$Cluster
    )
    Get-TervisClusterNode -cluster $Cluster |
    select -First 1 -ExpandProperty ADSite
}

function Get-TervisClusterSharedVolumeToStoreVMOSOn {
    param(
        [Parameter(Mandatory)][String]$Cluster
    )
    $EstimatedVMOSStorageSpace = 300
    $AmountOfSafetyEmptySpace = 300

    $CSVs = Get-TervisClusterSharedVolume -Cluster $Cluster 
    $CSVToStoreVMOS = $CSVs | 
    where metadata -NotContains "CX3-20" |
    where metadata -NotContains "Dedup" |
    where metadata -Contains "Auto-Tier" | 
    where {($_.FreeSpace - $EstimatedVMOSStorageSpace) -gt $AmountOfSafetyEmptySpace } | 
    sort FreeSpace -Descending | 
    Select -First 1

    $CSVToStoreVMOS
}

Function Get-TervisClusterSharedVolume {
    param(
        [Parameter(Mandatory)][String]$Cluster
    )

    process {
        $CSVs = Get-ClusterSharedVolume -cluster $Cluster
        $CSVs | Mixin-ClusterSharedVolume
        $CSVs
    }
}

filter Mixin-ClusterSharedVolume {
    $_ | Add-Member -MemberType ScriptProperty -Name UsedSpace -Value { ($this.SharedVolumeInfo.partition.usedspace)/1024/1024/1024 }
    $_ | Add-Member -MemberType ScriptProperty -Name FreeSpace -Value { ($this.SharedVolumeInfo.partition.Freespace)/1024/1024/1024 }
    $_ | Add-Member -MemberType ScriptProperty -Name PercentFree -Value { ($this.SharedVolumeInfo.partition.PercentFree) }
    $_ | Add-Member -MemberType ScriptProperty -Name Size -Value { ($this.SharedVolumeInfo.partition.Size)/1024/1024/1024 }
    $_ | Add-Member -MemberType ScriptProperty -Name MetaData -Value { 
        $($this.Name -split {$_ -eq "(" -or $_ -eq ")"})[1] -split ", "
    }
}
