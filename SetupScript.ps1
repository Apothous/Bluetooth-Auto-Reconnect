#-------------------------------------------------------------------------------------------------------#
# Script 1: Save Bluetooth device data to a CSV file or overwrite it if file already exists and 
# optionally create a windows scheduled task for the Bluetooth Auto Reconnect script
# Description:
#     This script searches the Windows Registry for all paired Bluetooth devices and exports their
#     names and MAC addresses to a CSV file and logs the event to a log file. It also gives the user
#     the option to automatically create a scheduled task to run the Bluetooth Auto Reconnect script
#     at user logon.
#-------------------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------------------#
# Function to check for Admin privileges
#-------------------------------------------------------------------------------------------------------#
function TestIsAdmin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#-------------------------------------------------------------------------------------------------------#
# Ensure script is running with Administrator privileges
#-------------------------------------------------------------------------------------------------------#
if (-not (TestIsAdmin)) {
    Write-Warning "Administrator privileges are required to run this script properly, especially for creating scheduled tasks."
    Write-Host "Attempting to re-launch with elevated privileges..."
    Start-Sleep 1
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File ""$($MyInvocation.MyCommand.Path)""" -Verb RunAs
    exit
}

# Script continues here if already running as Admin or after successful elevation

#-------------------------------------------------------------------------------------------------------#
# Static Global Variables
#-------------------------------------------------------------------------------------------------------#

# Define the path to the log file where script activity will be recorded
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "LogFile.txt"

# Define the path to the CSV file where Bluetooth device information will be saved
$deviceFile = Join-Path -Path $PSScriptRoot -ChildPath "BTDevice.csv"

#-------------------------------------------------------------------------------------------------------#
# Function to log activity
#-------------------------------------------------------------------------------------------------------#

# Function to write messages to the log file with a timestamp
function LogMessage {
    param([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -Append $logFile
}

#-------------------------------------------------------------------------------------------------------#
# Function to retrieve all currently paired bluetooth devices and assign numeric IDs
#-------------------------------------------------------------------------------------------------------#

function RetrieveBluetoothInfo {
    $counter = 1
    $bluetoothDevices = Get-ChildItem -Path "HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices" | ForEach-Object {
        $mac = $_.PSChildName
        $name = $_.GetValue("Name")

        if ($name) {
            $deviceName = ($name | ForEach-Object { [char]$_ }) -join ''

            # Create object with numeric DeviceID
            $device = [PSCustomObject]@{
                DeviceID    = $counter
                DeviceName  = $deviceName
                MACAddress  = $mac.ToUpper()
            }

            $counter++
            return $device
        }
    }
    return $bluetoothDevices
}

#-------------------------------------------------------------------------------------------------------#
# Function to get user input
#-------------------------------------------------------------------------------------------------------#
function GetUserInput {
    Write-Host "Please choose your desired device from the list using the Device ID number: " 
    $userInput = Read-Host
    return $userInput
}

#-------------------------------------------------------------------------------------------------------#
# Function to match user input to list of devices
#-------------------------------------------------------------------------------------------------------#

function MatchUserInput {
    $retrievedDevices = RetrieveBluetoothInfo

    while ($true) {
        $userChoice = GetUserInput
        foreach ($device in $retrievedDevices) {
            if ($device.DeviceID -eq [int]$userChoice) {
                return $device
            }
        }
        Write-Host "Invalid choice. Please try again."
        Write-Host ""
    }
}

#-------------------------------------------------------------------------------------------------------#
# Function to save the selected device to a CSV file
#-------------------------------------------------------------------------------------------------------#

function SaveToCSV {
    $chosenDevice = MatchUserInput

    # Export the list of devices to a CSV file, overwriting any existing file
    $chosenDevice | Select-Object DeviceName, MACAddress | Export-Csv -Path $deviceFile -NoTypeInformation -Force

    # Log that device information was saved
    LogMessage "New Bluetooth Device Information Saved to BTDevice.csv"

    # Display a confirmation message to the user
    Write-Host ""
    Write-Host "Bluetooth devices saved to $deviceFile"
    Write-Host ""
}

#-------------------------------------------------------------------------------------------------------#
# Function to create/update the Scheduled Task
#-------------------------------------------------------------------------------------------------------#
function RegisterBluetoothReconnectTask {
    if (-not (TestIsAdmin)) {
        Write-Host ""
        Write-Warning "Administrator privileges are required to create or update the scheduled task."
        Write-Warning "Please re-run this script as an Administrator if you wish to set up the scheduled task."
        return # This return might be unreachable
    }

    $taskName = "Bluetooth Auto Reconnect"
    # Using the more detailed description as per your preference
    $taskDescription = "This task continuously runs a PowerShell script called Bluetooth Auto Reconnect every minute. The script enables and disables the A2DP service for a paired Bluetooth audio device specified using the devices MAC Address. By toggling the A2DP service the PC sends a signal to connect with the paired Bluetooth audio device if the device is within range and available to connect."
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "BluetoothAutoReconnect.ps1"
    $workingDirectory = $PSScriptRoot

    # Action: What the task does
    $taskActionParams = @{
        Execute          = 'powershell.exe'
        Argument         = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
        WorkingDirectory = $workingDirectory
    }
    $action = New-ScheduledTaskAction @taskActionParams

    # Add the task triggers
    $triggers = @(
        New-ScheduledTaskTrigger -AtStartup
        New-ScheduledTaskTrigger -AtLogOn
        New-ScheduledTaskTrigger -AtWorkstationUnlock
    )

    # Principal: Who the task runs as
    $taskPrincipalParams = @{
        UserId    = "NT AUTHORITY\SYSTEM"
        LogonType = 'ServiceAccount'
        RunLevel  = 'Highest'
    }
    $principal = New-ScheduledTaskPrincipal @taskPrincipalParams

    # Settings: Additional task settings
    $taskSettingsParams = @{
        AllowStartIfOnBatteries     = $true
        DontStopIfGoingOnBatteries  = $true
        Compatibility               = 'Win8' # Corresponds to "Configure for: Windows 10"
        Hidden                      = $true
        ExecutionTimeLimit          = (New-TimeSpan -Seconds 0) # Indefinite
        MultipleInstances           = 'IgnoreNew'
        StartWhenAvailable          = $true
        RestartCount                = 3
        RestartInterval             = (New-TimeSpan -Minutes 1)
    }
    $settings = New-ScheduledTaskSettingsSet @taskSettingsParams

    Write-Host "Attempting to register scheduled task '$taskName'..."
    Start-Sleep 2

    try {
        # Unregister the task if it already exists, to ensure settings are updated
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Host "Any existing task named '$taskName' has been removed."
            Start-Sleep 2
        }
    }
    catch {
        Write-Warning "Could not remove existing task '$taskName' (it might not exist or an error occurred): $($_.Exception.Message)"
        Start-Sleep 3
    }

    try {
        $registerTaskParams = @{
            TaskName    = $taskName
            Action      = $action
            Trigger     = $triggers
            Principal   = $principal
            Settings    = $settings
            Description = $taskDescription
            Force       = $true
        }
        Register-ScheduledTask @registerTaskParams
        Write-Host "Scheduled task '$taskName' successfully registered with multiple triggers (Startup, Logon, Workstation Unlock)."
        LogMessage "Scheduled task '$taskName' registered/updated with multiple triggers."
        Start-Sleep 2
    }
    catch {
        Write-Error "Failed to register scheduled task '$taskName': $($_.Exception.Message)"
        LogMessage "Error registering scheduled task '$taskName': $($_.Exception.Message)"
        Start-Sleep 2
    }
}

#-------------------------------------------------------------------------------------------------------#
# Main Logic
#-------------------------------------------------------------------------------------------------------#

# Initial log entry for this script run
LogMessage "FindBTDeviceInfo.ps1 script started."

# Print the list of retrieved devices to the console
$retrievedDevices = RetrieveBluetoothInfo
$retrievedDevices | Select-Object DeviceID, DeviceName | Format-Table -AutoSize

SaveToCSV

# Ask user if they want to create/update the scheduled task
$choice = Read-Host "Do you want to create/update a scheduled task to run BluetoothAutoReconnect.ps1? (y/n)"
if ($choice -eq 'y' -or $choice -eq 'Y') {
    RegisterBluetoothReconnectTask
    Write-Host "Successfully registered scheduled task '$taskName'"
    Start-Sleep 3
} else {
    Write-Host "Skipping scheduled task creation."
    Start-Sleep 3
}
LogMessage "FindBTDeviceInfo.ps1 script finished."
