#Requires -Modules VMware.VimAutomation.Common

<#
.SYNOPSIS
This script will produce a CSV report of CPU/Memory recommendations from vROPs for vSphere VMs.

.DESCRIPTION
The logic for right sizing is as follows:
- Get current CPU/Memory allocations for a given VM
- Get CPU/Memory recommendation values from vROPs (these are aggressive)
- Determine how many % the current allocations for CPU/Memory deviate from the recommendation (regardless of undersized/oversized)
- If CPU and/or Memory are >= $RIGHT_SIZE_TOLERANCE_PERCENTAGE, VM will be right-sized
- Depending on whether CPU and/or Memory need to be resized, the following logic will determine what values to use:
    - Difference Value = Recommendation - Current Allocation
    - Recommended Value = Current Allocation + ($RIGHT_SIZE_BUFFER_PERCENTAGE * Difference Value)

The percentage values for tolerance/buffer can be modified in the Constants section

.PARAMETER vCenterServer
FQDN of vCenter Server to query (NOTE: Script will also query all Linked vCenter Servers)

.PARAMETER vCenterUsername
Username in UPN format (user@domain) for above vCenterServer

.PARAMETER vCenterPassword
Password for vCenterUsername

.PARAMETER allLinked
Boolean flag of whether or not to connect to all vCenter Servers connected to the specified one above in ELM

.PARAMETER vmFilter
--OPTIONAL--
Hashtable of VM filter values in the format @{Filter1 = Value1; Filter2 = Value2; Filter3 = "Value 3"}
See example for some suggestions on usage. The filters are flags that you would normally call with Get-VM, ie -Name, -Id, etc.

.PARAMETER vROPsServer
fqdn of vROPs Server that is configured to monitor VMs within vCenterServer + Linked VCs

.PARAMETER vROPsUsername
Username in UPN format (user@domain) for above vROPsServer. It is recommended to use the 'admin' user.

.PARAMETER vROPsPassword
Password for vROPsUsername

.PARAMETER CSVOutPath
--OPTIONAL--
Path to where the generated CSV file will be placed. If not specified, the csv file will be generated into the folder where the script resides.

.EXAMPLE
.\rightSize-kc-readOnly.ps1 `
-vCenterServer "vc-fqdn.local" `
-vCenterUsername "foo@bar.com" `
-vCenterPassword "p4ssw0rD1_" `
-allLinked $true `
-vmFilter @{Tag = "Tag Name"; Name = "*dev*"; Datastore = "DS Name"} `
-vROPsServer "vrops-fqdn.local" `
-vROPsUsername "admin" `
-vROPsPassword "p4ssw0rD1_1234"

.NOTES
- The operation to pull stats from vROPs can take a considerable amount of time
- There is a REST API authentication issue present in vROPs 6.6.0, resolved in 6.6.1. Please avoid using 6.6.0 with this script
#>

#Define Params
param
(
  [string   ]$vCenterServer,
  [string   ]$vCenterUsername,
  [string   ]$vCenterPassword,
  [boolean  ]$allLinked = $true,
  [string   ]$vSphereTag_RightSizeBooleanName,
  [string   ]$vSphereTag_RightSizeBooleanCategoryName,
  [string   ]$vSphereTag_GroupingCategoryName,
  [hashtable]$vmFilter = @{},
  [string   ]$vROPsServer,
  [string   ]$vROPsUsername,
  [string   ]$vROPsPassword,
  [string   ]$CSVOutPath = (Split-Path $script:MyInvocation.MyCommand.Path).toString()
)

#Define constants
New-Variable -Name 'MAX_RIGHT_SIZE_THREADS_ASYNC'           -Value  4    -Option Constant # How many concurrent right sizing operations to run together
New-Variable -Name 'NUM_DAYS_BACK_VROPS'                    -Value -5    -Option Constant # Number of days to look back at data in vROPs
New-Variable -Name 'NUM_MINUTES_TO_RUN_FOR'                 -Value  240  -Option Constant # 4 hours
New-Variable -Name 'NUM_SECONDS_TO_WAIT_FOR_TOOLS_SHUTDOWN' -Value  120  -Option Constant # 2 minutes
New-Variable -Name 'NUM_SECONDS_TO_WAIT_FOR_TOOLS_STARTUP'  -Value  120  -Option Constant # 2 minutes
New-Variable -Name 'RIGHT_SIZE_BUFFER_PERCENTAGE'           -Value  0.15 -Option Constant # Aggressive recommendation from vROPs + RIGHT_SIZE_BUFFER_PERCENTAGE%
New-Variable -Name 'RIGHT_SIZE_TOLERANCE_PERCENTAGE'        -Value  0.15 -Option Constant # How many % the VMs current values need to deviate away from the recommended value for in order to included in right sizing

#Define variables
$targetVMs   = @{} #Empty dictionary
$excludedVMs = @{} #Empty dictionary
$startTime   = Get-Date
$cutOffTime  = $startTime.AddMinutes($NUM_MINUTES_TO_RUN_FOR)

##########################################
# SETUP CONNECTIONS
##########################################
#Create vCenter credentials
$vCenterSecurePassword = ConvertTo-SecureString "$vCenterPassword" -AsPlainText -Force
$vCenterCredential     = New-Object System.Management.Automation.PSCredential ($vCenterUsername, $vCenterSecurePassword)

switch ($allLinked)
{
    $true  {
        Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Scope User -Confirm:$false
        $connectELM = @{AllLinked = $true}
    }
    $false {
        Set-PowerCLIConfiguration -DefaultVIServerMode Single   -Scope User -Confirm:$false
        $connectELM = @{}
    }
}

#Create VC connection
try {
    $vCenterConnection = Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential @connectELM
}
catch {
    Write-Host "Error creating VC Connection!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit -1
}

#Create vROPs credentials
$vROPsSecurePassword   = ConvertTo-SecureString "$vROPsPassword" -AsPlainText -Force
$vROPsCredential       = New-Object System.Management.Automation.PSCredential ($vROPsUsername, $vROPsSecurePassword)

# Create vROPs Connection
try {
    Connect-OMServer -Server $vROPsServer -Credential $vROPsCredential | Out-Null
}
catch {
    Write-Host "Error creating vROPs Connection!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit -1
}

<#########################################
# PULL LIST OF VMS TO BE RIGHT SIZED
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
    $targetVMs.Add($groupName, @())
    $excludedVMs.Add($groupName, @())

    # Filter all VMs for the current group name
    $vmsInGroup = $vmGroupTagAssignment | Where-Object {$_.Tag.Name -eq $groupName} | select entity;

    # Interate through the VMs in this group
    foreach ($vm in $vmsInGroup.Entity.Name)
    {
        # Check if the VM is to be resized
        if ($VMsToBeResized.Name -contains $vm)
        {
            $targetVMs.$groupName += $vm
        }
        # If we get to here it's either because the VM is tagged with the $vSphereTag_RightSizeBooleanName tag or simply doesn't have a tag from $vSphereTag_RightSizeBooleanCategoryName
        else
        {
            $excludedVMs.$groupName += $vm
        }
    }
}

##########################################
# SCRIPT BLOCK TO BE RUN IN ASYNC
##########################################
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
