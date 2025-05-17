#-------------------------------------------------------------------------------------------------------#
# Script 1: Save Bluetooth device data to a CSV file (overwrite if file exists)
# Description:
#     This script searches the Windows Registry for all paired Bluetooth devices and exports their
#     names and MAC addresses to a CSV file. It also logs the event to a log file.
#-------------------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------------------#
# User Configuration
#-------------------------------------------------------------------------------------------------------#

# Define the path to the log file where script activity will be recorded
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "LogFile.txt"

# Define the path to the CSV file where Bluetooth device information will be saved
$deviceFile = Join-Path -Path $PSScriptRoot -ChildPath "Devices.csv"

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
        Write-Host ""
    }
}


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
# Function to save to CSV file
#-------------------------------------------------------------------------------------------------------#

function SaveToCSV {
    $chosenDevice = MatchUserInput

    # Export the list of devices to a CSV file, overwriting any existing file
    $chosenDevice | Export-Csv -Path $deviceFile -NoTypeInformation -Force

    # Log that device information was saved
    Log-Message "New Bluetooth Device Information Saved to Devices.csv"

    # Display a confirmation message to the user
    Write-Host ""
    Write-Host "Bluetooth devices saved to $deviceFile"
}

#-------------------------------------------------------------------------------------------------------#
# Main Logic
#-------------------------------------------------------------------------------------------------------#

# Print the list of retrieved devices to the console
$retrievedDevices = RetrieveBluetoothInfo
$retrievedDevices | Format-Table -AutoSize

SaveToCSV
