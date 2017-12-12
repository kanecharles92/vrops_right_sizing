<#

.SYNOPSIS
This script will right-size and enable hot-add for virtual machines that contain the specified vSphere tag parameters.

.DESCRIPTION
This script will do the following:
- Get all VMs from vCenter(s) with the $vSphereTag_RightSizeBooleanName tag from the $vSphereTag_RightSizeBooleanCategoryName category
- Get a list of all tag names (used for grouping) from tag category $vSphereTag_GroupingCategoryName
- Based upon the VMs collected in the first query (those slated for right-sizing), categorise those VMs into groups based upon the second query (their associated group)
- Determine the order in which the groups are to be processed
- TODO: Keep populating

.EXAMPLE
./rightSize.ps1 `
-vCenterServer "vmvimg01a.afp.le" `
-vCenterUsername "sampleuser" `
-vCenterPassword "samplepassword" `
-vSphereTag_RightSizeBooleanName "RightSize-Yes" `
-vSphereTag_RightSizeBooleanCategoryName "Right-Sizing" `
-vSphereTag_GroupingCategoryName "SCCM-Collections" `
-vROPsServer "vmonitoring.afp.le" `
-vROPsUsername "sampleuser" `
-vROPsPassword "samplepassword"

.NOTES
- Prerequisite that vSphere tags/categories are created/applied to VMs in order to be selected/grouped by this script

#>


#TODO: Add some error handling around not importing the module properly, as well as not having a new enough version of PowerCLI to use ops manager cmdlets

#Define Params
param
(
  [string]$vCenterServer,
  [string]$vCenterUsername,
  [string]$vCenterPassword,
  [string]$vSphereTag_RightSizeBooleanName,
  [string]$vSphereTag_RightSizeBooleanCategoryName,
  [string]$vSphereTag_GroupingCategoryName,
  [string]$vROPsServer,
  [string]$vROPsUsername,
  [string]$vROPsPassword
)

#Define constants
New-Variable -Name 'RIGHT_SIZE_TOLERANCE_PERCENTAGE' -Value 0.25 -Option Constant # 25%
New-Variable -Name 'RIGHT_SIZE_BUFFER_PERCENTAGE' -Value 0.25 -Option Constant # aggressive recommendation from vROPs + 25%
#New-Variable -Name 'T-SHIRT_SIZE_TOLERANCE_PERCENTAGE' -Value 0.15 -Option Constant # 15%
New-Variable -Name 'NUM_MINUTES_TO_RUN_FOR' -Value 240 -Option Constant # 4 hours
New-Variable -Name 'NUM_SECONDS_TO_WAIT_FOR_TOOLS_SHUTDOWN' -Value 120 -Option Constant # 2 minutes
New-Variable -Name 'NUM_SECONDS_TO_WAIT_FOR_TOOLS_STARTUP' -Value 120 -Option Constant # 2 minutes
New-Variable -Name 'MAX_RIGHT_SIZE_THREADS_ASYNC' -Value 10 -Option Constant # How many concurrent right sizing operations to run together

#Define variables
$targetVMs = @{} #Empty dictionary
$excludedVMs = @{} #Empty dictionary
$startTime = Get-Date 
$cutOffTime = $startTime.AddMinutes($NUM_MINUTES_TO_RUN_FOR)

<#########################################
#
#SETUP CONNECTIONS
#
#########################################>

#Import Module
Import-Module VMware.PowerCLI

#Configure PowerCLI for multiple servers in linked mode
Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Scope User -Confirm:$false

#Create vCenter credentials
$vCenterSecurePassword = ConvertTo-SecureString "$vCenterPassword" -AsPlainText -Force
$vCenterCredential = New-Object System.Management.Automation.PSCredential ($vCenterUsername, $vCenterSecurePassword)

#Create vROPs credentials
$vROPsSecurePassword = ConvertTo-SecureString "$vROPsPassword" -AsPlainText -Force
$vROPsCredential = New-Object System.Management.Automation.PSCredential ($vROPsUsername, $vROPsSecurePassword)

#Make connection to vCenter
Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential -AllLinked

#Make connection to vROPs (try/catch required because vROPs API is displaying flaky behavior)
Try
{
    Connect-OMServer -Server $vROPsServer -Credential $vROPsCredential -ErrorAction SilentlyContinue
}
Catch
{
    while ($global:DefaultOMServers.Count -le 0)
    {
        # Attempt to connect again
        Connect-OMServer -Server $vROPsServer -Credential $vROPsCredential -ErrorAction SilentlyContinue

        # Sleep 2 seconds then retry
        sleep 2
    }
}

<#########################################
#
#PULL LIST OF VMS TO BE RIGHT SIZED
#
#########################################>

# Get all VMs that are grouped as per the chosen resizing grouping method
$vmGroupTagAssignment = Get-TagAssignment -Category (Get-TagCategory $vSphereTag_GroupingCategoryName)

# Get distinct list of group names
$groupNames = $vmGroupTagAssignment.tag.name | Sort-Object | Get-Unique

# Get all VMs across all vCenters to be resized
$VMsToBeResized = Get-VM -Tag (Get-Tag -Name $vSphereTag_RightSizeBooleanName -Category (Get-TagCategory -Name $vSphereTag_RightSizeBooleanCategoryName))

# Determine which VMs are going to be resized as per their groups/boolean values
# For each of the group names (pulled from the tag names in $vSphereTag_GroupingCategoryName
foreach ($groupName in $groupNames)
{
    # Setup the hashtables for this group
    $targetVMs.Add($groupName, @{})
    $excludedVMs.Add($groupName, @{})

    # Filter all VMs for the current group name
    $vmsInGroup = $vmGroupTagAssignment | Where-Object {$_.Tag.Name -eq $groupName} | select entity;

    # Interate through the VMs in this group
    foreach ($vm in $vmsInGroup.Entity.Name)
    {
        # Check if the VM is to be resized
        if ($VMsToBeResized.Name -contains $vm)
        {
            $targetVMs.$groupName.Add($vm, @{})
        }
        # If we get to here it's either because the VM is tagged with the $vSphereTag_RightSizeBooleanName tag or simply doesn't have a tag from $vSphereTag_RightSizeBooleanCategoryName
        else
        {
            $excludedVMs.$groupName.Add($vm, @{})
        }
    }
}

<#
foreach ($key in $targetVMs.Keys)
{
    Write-Host $key
    Write-Host $targetVMs.$key.Keys
}
exit 0
#>

<#
EXAMPLE FROM ABOVE:
RightSize3
DVEVMOTMD000
RightSize1
DVEVMOTDR098
RightSize2
DVEVMOTDR097 DVEVMOTKC000
#>


<#########################################
#
#GATHER SOME STATS
#
#########################################>

# Scriptblock to be run in async
$aSyncResize = {
    Param
    (
        [string]$vmNameToResize
    )
    
    # Define some temporary variables
    $vSphereVM = (Get-VM $vmNameToResize)
    [boolean]$CPUneedsResizing = $false
    [boolean]$MEMneedsResizing = $false
    #[boolean]$useTShirtSizing = $false
    [boolean]$isUndersized = $false
    [boolean]$isOversized = $false
    [int]$CPUsAdded = 0
    [int]$CPUsReclaimed = 0
    [int]$MEMAdded = 0
    [int]$MEMReclaimed = 0
    $statsHash = @{}

    # Get vROPs stats
    $vmStats = $vSphereVM | Get-OMResource -Server $omServer | Get-OMStat `
    -IntervalType Days `
    -IntervalCount 1 `
    -RollupType Latest `
    -From (Get-Date).AddDays(-1) `
    -Key (Get-OMStatKey "cpu|oversized","cpu|consumption-unit-count","cpu|size.recommendation","cpu|numberToAdd","cpu|numberToRemove","mem|oversized","mem|size.recommendation","mem|actual.capacity")

    # Due to vROPs returning duplicate entries for some keys, we have to grab the newest value for each key
    $vmStatKeys = $vmStats | select Key -Unique

    foreach ($statKey in $vmStatKeys)
    {
        # Add the statKey name as the key, and the latest value for the value
        $statsHash.Add($statKey, ($vmStats | ? {$_.Key -eq $statKey} | Sort-Object -Property Time -Descending | Select-Object Value -First 1))
    }

    <#
    Will end up with a hashtable looking like:
    |-----------------|-----------|
    | KEY             | VALUE     |
    |-----------------|-----------|
    | cpu|oversized   |  1        |
    | mem|oversized   |  1        |
    | cpu|numberToAdd | -2        |
    | etc....         | etc....   |
    |-----------------|-----------|
    #>

    # Do some conditionals

    # Is memory or cpu oversized?
    if (($statsHashItem.Item("cpu|oversized") -eq 1) -or ($statsHashItem.Item("cpu|oversized") -eq 1)) { $isOversized = $true }

    # Is memory or cpu undersized?
    if (($statsHashItem.Item("cpu|numberToRemove") -ge 1) -or ($statsHashItem.Item("mem|actual.capacity") -lt $statsHashItem.Item("mem|size.recommendation")) { $isUndersized = $true }

    # At this point we will determine if (provided the VM is oversized/undersized) the VM is more than the tolerance % away from the recommended values
    if (($isUndersized) -or ($isOversized))
    {
        $currentCPU = $statsHashItem.Item("cpu|consumption-unit-count")
        $recommendedCPU = $statsHashItem.Item("cpu|size.recommendation")
        $currentMem = $statsHashItem.Item("mem|actual.capacity")
        $recommendedMem = $statsHashItem.Item("mem|size.recommendation")

        # Determine how much % we are from the recommended values with the current allocation
        $cpuDifferencePercentage = ($recommendedCPU - $currentCPU) / $currentCPU
        $memDifferencePercentage = ($recommendedMem - $currentMem) / $currentMem

        # Convert percentages to positive numbers if they are negative (ie oversized)
        if ($cpuDifferencePercentage -lt 0) { $cpuDifferencePercentage = $cpuDifferencePercentage * -1}
        if ($memDifferencePercentage -lt 0) { $memDifferencePercentage = $memDifferencePercentage * -1}

        # Determine if our difference values are outside of the tolerances defined in the constants at the top of the script
        if ($cpuDifferencePercentage -gt $RIGHT_SIZE_TOLERANCE_PERCENTAGE) { $CPUneedsResizing = $true }
        if ($memDifferencePercentage -gt $RIGHT_SIZE_TOLERANCE_PERCENTAGE) { $MEMneedsResizing = $true }

        # If either CPU or Memory need resizing
        if (($CPUneedsResizing) -or ($MEMneedsResizing))
        {
            # Define temp variables
            $VMShutdownCounter = 0

            <#
            - Gracefully shutdown VM
            - Wait x seconds for VM to shutdown
            - If not shutdown within x seconds, do force poweroff
            - Once shutdown, update CPU/Memory/Both to reflect recommended values
            - Power VM back on
            - Wait x seconds for VMware tools to be up
            - If not up after x seconds, make note that tools was not responding
            - Finish and move onto next VM
            #>

            # Do the resize work   
            # Gracefully shutdown VM if tools is running
            if (((($vSphereVM | Get-View).Guest.ToolsStatus) -eq "toolsOk") -or ((($vSphereVM | Get-View).Guest.ToolsStatus) -eq "toolsOld"))
            {
                $vSphereVM | Shutdown-VMGuest -Confirm:$false
                do
                {
                    # Sleep for 2 seconds between checks
                    sleep 2
                    # Increase counter by 2 seconds
                    $VMShutdownCounter += 2
                }
                while (((Get-VM $vSphereVM).PowerState -ne "PoweredOff") -and ($VMShutdownCounter -le $NUM_SECONDS_TO_WAIT_FOR_TOOLS_SHUTDOWN))

                # TODO: If we reach the NUM_SECONDS.... value, log that the VM didn't shutdown in time, do a force poweroff
            }
            # Otherwise force power down
            else
            {
                # Power off VM
                $vSphereVM | Stop-VM -Confirm:$false
                do
                {
                    # Sleep for 2 seconds between checks
                    sleep 2
                    # Increase counter by 2 seconds
                    $VMShutdownCounter += 2
                }
                while (((Get-VM $vSphereVM).PowerState -ne "PoweredOff") -and ($VMShutdownCounter -le $NUM_SECONDS_TO_WAIT_FOR_TOOLS_SHUTDOWN))
            }

            # By this point we are assuming that the VM was shutdown/powered off successfully

            # Apply CPU modification is necessary
            if ($CPUneedsResizing)
            {
                # Determine how many CPUs we want
                # Get the recommendation from vROPs 
                $recommendedCPU = $statsHashItem.Item("cpu|size.recommendation")

                # Add the buffer overhead so that we aren't using the aggressive figure
                # Force to be an integer so that the extra % doesn't give us a decimal
                $recommendedCPU = $recommendedCPU * (1 + $RIGHT_SIZE_BUFFER_PERCENTAGE)

                # Convert to int, ie round up/down to nearest whole number
                $recommendedCPU = [int]$recommendedCPU

                # Check to see if the recommendedCPU is an odd number, if it is we will be adding 1 so that we are always using even CPU figures
                if([bool]!($recommendedCPU%2))
                {
                    $recommendedCPU += 1
                }

                # Apply the recommended CPU value
                $vSphereVM | Set-VM -NumCpu $recommendedCPU -Confirm:$false
            }

            # Apply Memory modification is necessary
            if ($MEMneedsResizing)
            {
                # Determine how many MB of memory we want
                # Get the recommendation from vROPs 
                [decimal]$recommendedMem = $statsHashItem.Item("mem|size.recommendation")

                # Add the buffer overhead so that we aren't using the aggressive figure
                $recommendedMem = $recommendedMem * (1 + $RIGHT_SIZE_BUFFER_PERCENTAGE)

                # Ensure we convert the recommendedMem to an integer, ie nearest whole megabyte
                $recommendedMem = [int]$recommendedMem

                # Apply the recommended Memory value
                $vSphereVM | Set-VM -MemoryMB $recommendedMem -Confirm:$false
            }

            # Power on VM, wait for tools (use Wait-Tools cmdlet)
            $toolStartupCounter = 0
            $vSphereVM | Start-VM -Confirm:$false
            do
            {
                sleep 2
                $toolStartupCounter += 2
            }
            while ((($vSphereVM | Get-View).Guest.ToolsRunningStatus -ne "guestToolsRunning") -and ($toolStartupCounter -le $NUM_SECONDS_TO_WAIT_FOR_TOOLS_STARTUP))
        }
    }

    # TODO: Generate some output to a CSV file with fields like:
    # VMname | originalCPU | newCPU | CPUsReclaimed | CPUsAdded | | originalMEM | newMEM | MEMReclaimed | MEMAdded

    # TODO: Add/subtract from global CPUsAdded/CPUsRemoved, MEMAdded/MEMRemoved

    return $statsHash
}

# Loop through group names
foreach ($groupName in $groupNames)
{
    # Setup Runspace for this particular group
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MAX_RIGHT_SIZE_THREADS_ASYNC)
    $runspacePool.Open()
    $jobs = @()

    # For each VM in the group
    foreach($vm in $targetVMs.$groupName.Keys)
    {
        # Check to see if the current time is less than the end time
        if ((Get-Date) -lt $cutOffTime)
        {
            # Start a runspace job for this VM
            $job = [powershell]::Create().AddScript($aSyncResize).AddArgument($vm)
            $job.RunspacePool = $runspacePool
            $jobs += New-Object PSObject -Property @{
                vmName = $vm
                Pipe = $job
                Result = $job.BeginInvoke()
            }
        }
        else # We have run out of time
        {
            Write-Host "NO TIME REMAINING, EXITING"
            exit -1
        }
        
    }

    Write-Host $groupName - "Waiting.." -NoNewline

    Do 
    {
       Write-Host "." -NoNewline
       Start-Sleep -Seconds 1
    } 
    While ( $jobs.Result.IsCompleted -contains $false)
    
    Write-Host "All jobs completed!"
 
    $results = @()
    ForEach ($job in $jobs)
    {
        $results += $job.Pipe.EndInvoke($job.Result)
    }

    Write-Host $results
}