#-------------------------------------------------------------------------------------------------------#
# Script 2: Toggle Audio Service to connect PC to the Bluetooth Device
# Description:
#     This script disables and re-enables the A2DP (audio) service for a specific Bluetooth device.
#-------------------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------------------#
# User Configuration
#-------------------------------------------------------------------------------------------------------#

# Set the main loop interval
$interval = 30

# Set wait time after first noticing disconnected
$firstCheckInterval = 60

#-------------------------------------------------------------------------------------------------------#
# Static Global Variables
#-------------------------------------------------------------------------------------------------------#

# File path where logs will be saved
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "LogFile.txt"

# Define the path to the CSV file where Bluetooth device information will be saved
$deviceFile = Join-Path -Path $PSScriptRoot -ChildPath "BTDevice.csv"

# Define the A2DP profile service GUID (this is standard and should not change)
$a2dpGuid = [Guid]::Parse("0000110b-0000-1000-8000-00805f9b34fb")

#-------------------------------------------------------------------------------------------------------#
# Type Definition for Bluetooth Device Structure and Service Manager
#-------------------------------------------------------------------------------------------------------#

# Define the necessary C# structures and import Bluetooth function from Windows API (if not already loaded)
if (-not ([System.Management.Automation.PSTypeName]'BLUETOOTH_DEVICE_INFO').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    // This structure represents a Bluetooth device's key info
    [StructLayout(LayoutKind.Sequential)]
    public struct BLUETOOTH_DEVICE_INFO {
        public uint dwSize;           // Size of this structure in bytes (required by API)
        public ulong Address;         // Bluetooth address (converted from MAC string)
        public bool fConnected;       // Whether the device is currently connected
    }

    // Provides access to the BluetoothSetServiceState WinAPI function
    public class BtServiceManager {
        [DllImport("bthprops.cpl", CharSet = CharSet.Auto)]
        public static extern uint BluetoothSetServiceState(
            IntPtr hRadio,                          // Not used (set to zero)
            ref BLUETOOTH_DEVICE_INFO deviceInfo,   // Info struct for target device
            ref Guid guidService,                   // Service GUID (e.g., A2DP)
            uint dwServiceFlags                     // 0 = disable, 1 = enable
        );
    }
"@ -Language CSharp
}

#-------------------------------------------------------------------------------------------------------#
# Logs a timestamped message to the log file
#-------------------------------------------------------------------------------------------------------#
function LogMessage {
    param([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -Append $logFile
}

#-------------------------------------------------------------------------------------------------------#
# Converts a MAC address string (without colons) to a 64-bit unsigned integer
#-------------------------------------------------------------------------------------------------------#
function Convert-MacToULong {
    param($mac)
    return [Convert]::ToUInt64($mac, 16)
}

#-------------------------------------------------------------------------------------------------------#
# Function to check device status
#-------------------------------------------------------------------------------------------------------#
function DeviceStatus {
    # Load device configuration if not already loaded
    if (-not $script:deviceInfoLoaded) {
        LogMessage "Attempting to load device information from $deviceFile..."
        if (-not (Test-Path $deviceFile)) {
            $msg = "Device data file not found. Please run FindBTDeviceInfo.ps1 first to create it."
            Write-Error $msg
            LogMessage $msg
            throw $msg # Terminate script if config file is missing
        }

        $csvData = Import-Csv -Path $deviceFile
        if (-not $csvData -or $csvData.Count -eq 0) {
            $msg = "No device data found in $deviceFile or the file is empty. Please run FindBTDeviceInfo.ps1."
            Write-Error $msg
            LogMessage $msg
            throw $msg # Terminate script if config file is empty
        }

        $script:friendlyName = $csvData[0].DeviceName
        $script:deviceMAC = $csvData[0].MACAddress
        
        # Prepare Device Info Structure (moved here, populates script-scoped $script:info)
        $script:btAddr = Convert-MacToULong $script:deviceMAC
        $script:info = New-Object "BLUETOOTH_DEVICE_INFO"
        $script:info.dwSize = [System.Runtime.InteropServices.Marshal]::SizeOf([BLUETOOTH_DEVICE_INFO]) # More robust way to get size
        $script:info.Address = $script:btAddr
        $script:info.fConnected = $false # Connection status isn't required for toggling service here

        $script:deviceInfoLoaded = $true
        LogMessage "Device information loaded: Name='$($script:friendlyName)', MAC='$($script:deviceMAC)'."
        Write-Host "Device information loaded: Name='$($script:friendlyName)', MAC='$($script:deviceMAC)'."
    }

    $nameToCheck = $script:friendlyName # Use the loaded friendly name
    $statusOK = "OK"
    $foundAny = $false

    # Get all audio endpoint devices (connected or not)
    $devices = Get-PnpDevice -Class AudioEndpoint
    Write-Host "-------------------------------------------------------------------------"

    foreach ($device in $devices) {
        if ($device.FriendlyName -match "\(([^)]+)\)") {
            $nameInParens = $matches[1]
            if ($nameInParens -eq $nameToCheck) {
                $foundAny = $true
                Write-Host "Checking device: $($device.FriendlyName)"
                Write-Host "Status: $($device.Status)"
                if ($device.Status -eq $statusOK) {
                    Write-Host "Device is connected!"
                    return $true
                } else {
                    Write-Host "Device is disconnected!"
                    Write-Host ""
                }
            }
        }
    }

    if ($foundAny) {
        Write-Host "No matching '$nameToCheck' devices are currently connected."
    } else {
        Write-Host "No matching Bluetooth ($nameToCheck) devices are paired."
    }

    return $false
}

#-------------------------------------------------------------------------------------------------------#
# Main Execution: Toggle A2DP Service
#-------------------------------------------------------------------------------------------------------#

$firstCheck = $true

while ($true) {
    $connected = Device-Status
    if ($connected -eq $false) {
        if ($firstCheck -eq $true) {
            $firstCheck = $false
            Write-Host "Device has been disconnected!"
            Write-Host "Waiting $firstCheckInterval seconds before reconnecting."
            Start-Sleep $firstCheckInterval
        } else {
            try {
                Write-Host "Toggling A2DP service for $script:friendlyName (MAC: $script:deviceMAC)"
                # First, disable the A2DP service for the Bluetooth device
                [void][BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$a2dpGuid, 0)

                # Then, re-enable the A2DP service
                [void][BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$a2dpGuid, 1)
            }
            catch {
                # Capture and log any errors that occur during the toggle process
                Log-Message "Error occurred: $_"
            }
         }  
    } else {
        if ($firstCheck -eq $false) {
            $firstCheck = $true
        } 
    }
    Start-Sleep $interval
}
