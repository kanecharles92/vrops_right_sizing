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
-vROPsServer "vrops-fqdn.local" `
-vROPsUsername "admin" `
-vROPsPassword "p4ssw0rD1_1234"

.NOTES
The operation to pull stats from vROPs can take a considerable amount of time

#>

#Define Params
param
(
  [string]$vCenterServer,
  [string]$vCenterUsername,
  [string]$vCenterPassword,
  [boolean]$allLinked = $true,
  [string]$vROPsServer,
  [string]$vROPsUsername,
  [string]$vROPsPassword,
  [string]$CSVOutPath = (Split-Path $script:MyInvocation.MyCommand.Path).toString()
)

<#
###############
# TODO:
###############
- Add error handling
    - VM doesn't exist in vROPs
    - vROPs doesn't connect within x seconds or x attempts
    - "" for vCenter
    - Can't write to CSV file
    - Write XLS (with colors, headers, totals, etc) instead of CSV
#>

<#########################################
#
# IMPORT MODULES
#
#########################################>
Import-Module VMware.VimAutomation.Common

<#########################################
#
# DEFINE CONSTANTS
#
#########################################>
New-Variable -Name 'RIGHT_SIZE_TOLERANCE_PERCENTAGE' -Value   0.15 -Option Constant # How many % the VMs current values need to deviate away from the recommended value for in order to included in right sizing
New-Variable -Name 'RIGHT_SIZE_BUFFER_PERCENTAGE'    -Value   0.15 -Option Constant # Aggressive recommendation from vROPs + RIGHT_SIZE_BUFFER_PERCENTAGE%
New-Variable -Name 'MAX_RIGHT_SIZE_THREADS_ASYNC'    -Value   4    -Option Constant # How many concurrent right sizing operations to run together
New-Variable -Name 'NUM_DAYS_BACK_VROPS'             -Value  -5    -Option Constant # Number of days to look back at data in vROPs

<#########################################
#
# DEFINE/DECLARE VARS
#
#########################################>
$startTime = Get-Date

# Setup CSV Files
$csvFile = $CSVOutPath + "\rightSizeReport-" + $startTime.toString("dd.MM.yyyy_hh.mm.ss.tt") + ".csv"
$csvContents = @()

<#########################################
#
# SETUP CONNECTIONS
#
#########################################>
#Clear all current connections
Disconnect-VIServer * -Confirm:$false
Disconnect-OMServer * -Confirm:$false

#Create vCenter credentials
$vCenterSecurePassword = ConvertTo-SecureString "$vCenterPassword" -AsPlainText -Force
$vCenterCredential     = New-Object System.Management.Automation.PSCredential ($vCenterUsername, $vCenterSecurePassword)

#If we want to connect to all ELM connected VCs
if ($allLinked)
{
    #Configure PowerCLI for multiple servers in linked mode
    Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Scope User -Confirm:$false

    #Create VC connection
    $vCenterConnection = Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential -AllLinked
}
#We dont' want to connect to all ELM connected VCs
else
{
    #Configure PowerCLI for single server
    Set-PowerCLIConfiguration -DefaultVIServerMode Single -Scope User -Confirm:$false

    #Create VC connection
    $vCenterConnection = Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential
}

<#########################################
#
# PULL LIST OF VMS TO BE RIGHT SIZED
#
#########################################>
# Get all VMs
$allVMs = Get-VM

<#########################################
#
# SETUP ASYNC PREREQS
#
#########################################>

# Setup runspace
$runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MAX_RIGHT_SIZE_THREADS_ASYNC)
$runspacePool.Open()

# Setup arrays for our jobs
$runspaceJobs       = @()
$powerShellInstance = @()

$numOfVMs = $allVMs.Length

<#########################################
#
# SCRIPT BLOCK TO BE RUN IN ASYNC
#
#########################################>

# Scriptblock to be run in async
$aSyncResize = {
    Param
    (
        $vmToAnalyze,
        $csvContents,
        $vROPsServer,
        $vROPsUsername,
        $vROPsPassword,
        $NUM_DAYS_BACK_VROPS,
        $RIGHT_SIZE_TOLERANCE_PERCENTAGE,
        $RIGHT_SIZE_BUFFER_PERCENTAGE
    )

    $vmToAnalyze = $vmToAnalyze.Value

    #Create vROPs credentials
    $vROPsSecurePassword   = ConvertTo-SecureString "$vROPsPassword" -AsPlainText -Force
    $vROPsCredential       = New-Object System.Management.Automation.PSCredential ($vROPsUsername, $vROPsSecurePassword)

    #Make connection to vROPs (try/catch required because vROPs API is displaying flaky behavior)
    do
    {
        # Attempt to connect again
        Connect-OMServer -Server $vROPsServer -Credential $vROPsCredential | Out-Null

        sleep 2
    }
    while ($global:DefaultOMServers.Count -le 0)

    <#########################################
    #
    # GATHER SOME STATS
    #
    #########################################>
    # Retrieve stats
    $vmStats = $vmToAnalyze | `
        Get-OMResource -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | `
        Get-OMStat -ErrorAction SilentlyContinue -WarningAction SilentlyContinue `
        -IntervalType Days `
        -IntervalCount 1 `
        -RollupType Latest `
        -From (Get-Date).AddDays($NUM_DAYS_BACK_VROPS) `
        -Key (Get-OMStatKey "cpu|size.recommendation","mem|size.recommendation")

    # Define some temporary variables
    [boolean]$CPUneedsResizing = $false
    [boolean]$MEMneedsResizing = $false
    [boolean]$isUndersized     = $false
    [boolean]$isOversized      = $false
    [string]$notes = ""

    # # Prepare CSV Row
    $row = New-Object System.Object

    $currentCPU      = [int]$vmToAnalyze.NumCpu
    $recommendedCPU  = [int]($vmStats | ? {($_.Key -eq "cpu|size.recommendation") -and ($_.Resource.Name -eq $vmToAnalyze.Name)} | Sort-Object -Property Time -Descending | Select-Object Value -First 1).Value
    $currentMem      = [decimal]$vmToAnalyze.MemoryMB
    $recommendedMem  = [decimal]($vmStats | ? {($_.Key -eq "cpu|mem.recommendation") -and ($_.Resource.Name -eq $vmToAnalyze.Name)} | Sort-Object -Property Time -Descending | Select-Object Value -First 1).Value

    #Error handling in the event that vROPs has returned nothing for this VM, ie no data available
    if ([int]$recommendedCPU -eq [int]0)
    {
        $recommendedCPU = $currentCPU
        $notes += "CPU recommendation was invalid. "
    }
    if ([decimal]$recommendedMem -eq [decimal]0)
    {
        $recommendedMem = $currentMem
        $notes += "Memory recommendation was invalid. "
    }

    # Do some conditionals
    # Is memory or cpu oversized?
    if (($currentCPU -gt $recommendedCPU) -or ($currentMem -gt $recommendedMem)) { $isOversized = $true }

    # Is memory or cpu undersized?
    if (($currentCPU -lt $recommendedCPU) -or ($currentMem -lt $recommendedMem)) { $isUndersized = $true }

    # At this point we will determine if (provided the VM is oversized/undersized) the VM is more than the tolerance % away from the recommended values
    if (($isUndersized) -or ($isOversized))
    {
        # Determine the difference between current values and recommended values, regardless of whether under/over
        $cpuDifferenceValue = $recommendedCPU - $currentCPU
        $memDifferenceValue = $recommendedMem - $currentMem

        # Convert negative values to positive if necessary
        if ($cpuDifferenceValue -lt 0) { $cpuDifferenceValue = $cpuDifferenceValue * -1}
        if ($memDifferenceValue -lt 0) { $memDifferenceValue = $memDifferenceValue * -1}

        # Determine how much % we are from the recommended values with the current allocation
        $cpuDifferencePercentage = $cpuDifferenceValue / $currentCPU
        $memDifferencePercentage = $memDifferenceValue / $currentMem

        # Convert percentages to positive numbers if they are negative (ie oversized)
        if ($cpuDifferencePercentage -lt 0) { $cpuDifferencePercentage = $cpuDifferencePercentage * -1}
        if ($memDifferencePercentage -lt 0) { $memDifferencePercentage = $memDifferencePercentage * -1}

        # Determine if our difference values are outside of the tolerances defined in the constants at the top of the script
        if ($cpuDifferencePercentage -gt $RIGHT_SIZE_TOLERANCE_PERCENTAGE) { $CPUneedsResizing = $true }
        if ($memDifferencePercentage -gt $RIGHT_SIZE_TOLERANCE_PERCENTAGE) { $MEMneedsResizing = $true }

        # If either CPU or Memory need resizing
        if (($CPUneedsResizing) -or ($MEMneedsResizing))
        {
            # Apply CPU modification is necessary
            if ($CPUneedsResizing)
            {
                # Add the buffer overhead so that we aren't using the aggressive figure
                # Force to be an integer so that the extra % doesn't give us a decimal
                [decimal]$recommendedCPU = ($recommendedCPU + ($RIGHT_SIZE_BUFFER_PERCENTAGE * $cpuDifferenceValue))

                # Convert to int, ie round up/down to nearest whole number
                $recommendedCPU = [int]$recommendedCPU

                # Check to see if the recommendedCPU is an odd number, if it is we will be adding 1 so that we are always using even CPU figures
                if (![bool]!($recommendedCPU%2))
                {
                    $recommendedCPU += 1
                }

                # If the VM has been powered off/vROPs recommendation is 0, don't change the recommended value. ie don't modify CPU
                if ($recommendedCPU -eq 0)
                {
                    $recommendedCPU = $currentCPU
                }
            }

            # Apply Memory modification is necessary
            if ($MEMneedsResizing)
            {
                # Determine how many MB of memory we want
                # Add the buffer overhead so that we aren't using the aggressive figure
                [decimal]$recommendedMem = ($recommendedMem + ($RIGHT_SIZE_BUFFER_PERCENTAGE * $memDifferenceValue))

                # Ensure we convert the recommendedMem to an integer, ie nearest whole megabyte
                $recommendedMem = [int]$recommendedMem

                # If the VM has been powered off/vROPs recommendation is 0, don't change the recommended value. ie don't modify CPU
                if ($recommendedMem -eq 0)
                {
                    $recommendedMem = $currentMem
                }
            }
        }
    }

    # Add content to be put in the CSV report
    $row | Add-Member -MemberType NoteProperty -Name "VM|Name"                       -Value $vmToAnalyze.Name
    $row | Add-Member -MemberType NoteProperty -Name "vCPU|Original Count"           -Value $currentCPU
    $row | Add-Member -MemberType NoteProperty -Name "vCPU|New CPU Count"            -Value $recommendedCPU
    $row | Add-Member -MemberType NoteProperty -Name "vCPU|Num Added"                -Value ($recommendedCPU - $currentCPU)
    $row | Add-Member -MemberType NoteProperty -Name "vCPU|Num Reclaimed"            -Value ($currentCPU - $recommendedCPU)
    $row | Add-Member -MemberType NoteProperty -Name "vMEM|Original Allocation (MB)" -Value $currentMem
    $row | Add-Member -MemberType NoteProperty -Name "vMEM|New Allocation (MB)"      -Value $recommendedMem
    $row | Add-Member -MemberType NoteProperty -Name "vMEM|Added (MB)"               -Value ($recommendedMem - $currentMem)
    $row | Add-Member -MemberType NoteProperty -Name "vMEM|Reclaimed (MB)"           -Value ($currentMem - $recommendedMem)
    $row | Add-Member -MemberType NoteProperty -Name "VM|Notes"                      -Value $notes

    $csvContents.Value += $row

    return $row
}

<#########################################
#
# QUEUE UP ASYNC JOBS
#
#########################################>

# Add a job for every VM
for ($i = 0; $i -lt $numOfVMs; $i++)
{
    $powerShellInstance += [powershell]::create()
    $powerShellInstance[$i].RunspacePool = $runspacePool

    # Start a runspace job for this VM
    [void]$powerShellInstance[$i].AddScript($aSyncResize)
    $powerShellInstance[$i].AddArgument(([REF]$allVMs[$i]))
    $powerShellInstance[$i].AddArgument(([REF]$csvContents))
    $powerShellInstance[$i].AddArgument($vROPsServer)
    $powerShellInstance[$i].AddArgument($vROPsUsername)
    $powerShellInstance[$i].AddArgument($vROPsPassword)
    $powerShellInstance[$i].AddArgument($NUM_DAYS_BACK_VROPS)
    $powerShellInstance[$i].AddArgument($RIGHT_SIZE_TOLERANCE_PERCENTAGE)
    $powerShellInstance[$i].AddArgument($RIGHT_SIZE_BUFFER_PERCENTAGE)

    $runspaceJobs += $powerShellInstance[$i].BeginInvoke()
}

<#########################################
#
# WAIT FOR JOBS TO FINISH
#
#########################################>

# Wait for jobs to be finished, new jobs will be added when queue drains of runspace jobs
for ($i = 0; $i -lt $allVMs.Length; $i++) {
    try {
        $CurrentThreads = if ($MAX_RIGHT_SIZE_THREADS_ASYNC -gt ($allVMs.Length-$i)) {$allVMs.Length-$i} else {$MAX_RIGHT_SIZE_THREADS_ASYNC}
        $ProgressSplat = @{
            Activity         = "Running Query"
            Status           = "Starting threads"
            CurrentOperation = "$($allVMs.Length) threads created - $CurrentThreads threads concurrently running - $($allVMs.Length-$i) threads open"
            PercentComplete  = $i / $allVMs.Length * 100
        }
        Write-Progress @ProgressSplat

        if ($i -lt $numOfVMs)
        {
            $currentVMNumber = $i + 1
            Write-Host "Processing VM ($currentVMNumber/$numOfVMs):" $allVMs[$i]
        }

        $powerShellInstance[$i].EndInvoke($runspaceJobs[$i])

    } catch {
        write-warning "error: $_"
    }
}

<#########################################
#
# WRITE TO CSV
#
#########################################>

# Write the CSV file out
$csvContents | Export-CSV -NoTypeInformation -Path $csvFile

Write-Host "Output File: "$csvFile