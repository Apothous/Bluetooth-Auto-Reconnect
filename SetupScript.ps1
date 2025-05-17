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

        Write-Host ""
        Write-Host "Invalid choice. Please try again."
        Write-Host "----------------------------------------------------------------------------------------"
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
    Write-Host "----------------------------------------------------------------------------------------"
}

#-------------------------------------------------------------------------------------------------------#
# Function to check for Admin privileges
#-------------------------------------------------------------------------------------------------------#
function Test-IsAdmin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#-------------------------------------------------------------------------------------------------------#
# Function to create/update the Scheduled Task
#-------------------------------------------------------------------------------------------------------#
function Register-BluetoothReconnectTask {
    if (-not (Test-IsAdmin)) {
        Write-Warning "Administrator privileges are required to create or update the scheduled task."
        Write-Warning "Please re-run this script as an Administrator if you wish to set up the scheduled task."
        return
    }

    $taskName = "Bluetooth Auto Reconnect"
    $taskDescription = "This task continuously runs a PowerShell script called Bluetooth Auto Reconnect every minute. 
    The script enables and disables the A2DP service for a paired Bluetooth audio device specified using the devices MAC Address. 
    By toggling the A2DP service the PC sends a signal to connect with the paired Bluetooth audio device if the device is within 
    range and available to connect."

    $scriptPath = Join-Path 
        -Path $PSScriptRoot 
        -ChildPath "BluetoothAutoReconnect.ps1"
    $workingDirectory = $PSScriptRoot

    # Action: What the task does
    $action = New-ScheduledTaskAction 
        -Execute 'powershell.exe' 
        -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" 
        -WorkingDirectory $workingDirectory

    # Trigger: When the task runs
    $trigger = New-ScheduledTaskTrigger 
        -AtStartup
    $trigger = New-ScheduledTaskTrigger 
        -AtLogOn

    # Principal: Who the task runs as
    $principal = New-ScheduledTaskPrincipal 
        -UserId "NT AUTHORITY\SYSTEM" 
        -LogonType ServiceAccount 
        -RunLevel Highest 

    # Settings: Additional task settings
    $settings = New-ScheduledTaskSettingsSet 
        -AllowStartIfOnBatteries 
        -DontStopIfGoingOnBatteries `
        -Compatibility Win8 ` # Corresponds to "Configure for: Windows 10" (Win8 is a common mapping)
        -Hidden $true ` # Hidden" checkbox
        -ExecutionTimeLimit (New-TimeSpan -Seconds 0) ` # Unchecked "Stop the task if it runs longer than" (0 = indefinite)
        -MultipleInstances IgnoreNew ` # "Do not start a new instance"
        -StartWhenAvailable ` # "Run task as soon as possible after a scheduled restart is missed"
        -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) # "If the task fails, restart every 1 minute, 3 times"

    Write-Host "Attempting to register scheduled task '$taskName'..."

    try {
        # Unregister the task if it already exists, to ensure settings are updated
        Get-ScheduledTask 
            -TaskName $taskName 
            -ErrorAction SilentlyContinue | Unregister-ScheduledTask 
            -Confirm:$false
        Write-Host "Any existing task named '$taskName' has been removed."
    }
    catch {
        Write-Warning "Could not remove existing task '$taskName' (it might not exist or an error occurred): $($_.Exception.Message)"
    }

    try {
        Register-ScheduledTask 
            -TaskName $taskName 
            -Action $action 
            -Trigger $trigger 
            -Principal $principal 
            -Settings $settings 
            -Description $taskDescription 
            -Force
        Write-Host "Scheduled task '$taskName' successfully registered to run at user logon."
        LogMessage "Scheduled task '$taskName' registered/updated."
    }
    catch {
        Write-Error "Failed to register scheduled task '$taskName': $($_.Exception.Message)"
        LogMessage "Error registering scheduled task '$taskName': $($_.Exception.Message)"
    }
    Write-Host "----------------------------------------------------------------------------------------"
}

#-------------------------------------------------------------------------------------------------------#
# Main Logic
#-------------------------------------------------------------------------------------------------------#

# Initial log entry for this script run
LogMessage "FindBTDeviceInfo.ps1 script started."

# Check if running with Admin privileges early on, as it's needed for scheduled task creation
if (-not (Test-IsAdmin)) {
    Write-Warning "This script is not running with Administrator privileges."
    Write-Warning "You will be able to select and save a device, but creating/updating the scheduled task will be skipped."
}

# Print the list of retrieved devices to the console
$retrievedDevices = RetrieveBluetoothInfo
$retrievedDevices | Select-Object DeviceID, DeviceName | Format-Table -AutoSize

SaveToCSV

# Ask user if they want to create/update the scheduled task
Write-Host ""
$choice = Read-Host "Do you want to create/update a scheduled task to run BluetoothAutoReconnect.ps1? (y/n)"
if ($choice -eq 'y' -or $choice -eq 'Y') {
    Register-BluetoothReconnectTask
} else {
    Write-Host "Skipping scheduled task creation."
    Write-Host "----------------------------------------------------------------------------------------"
}
LogMessage "FindBTDeviceInfo.ps1 script finished."
