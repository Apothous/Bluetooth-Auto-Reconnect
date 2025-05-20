#-------------------------------------------------------------------------------------------------------#
# Script 2: Toggle Audio Service to connect PC to the Bluetooth Device
# Description:
#     This script disables and re-enables the A2DP (audio) service for a specific Bluetooth device.
#-------------------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------------------#
# User Configuration
#-------------------------------------------------------------------------------------------------------#

# Set the main loop interval
$script:interval = 30

# Set wait time after first noticing disconnected
$script:firstCheckInterval = 60

#-------------------------------------------------------------------------------------------------------#
# Static Global Variables
#-------------------------------------------------------------------------------------------------------#

# File path where logs will be saved
$script:logFile = Join-Path -Path $PSScriptRoot -ChildPath "LogFile.txt"

# Define the path to the CSV file where Bluetooth device information will be saved
$script:deviceFile = Join-Path -Path $PSScriptRoot -ChildPath "BTDevice.csv"

# Define the A2DP profile service GUID (this is standard and should not change)
$script:a2dpGuid = [Guid]::Parse("0000110b-0000-1000-8000-00805f9b34fb")


$script:deviceData =$null
$script:info = $null

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
    "$timestamp - $message" | Out-File -Append $script:logFile
}

#-------------------------------------------------------------------------------------------------------#
# Converts a MAC address string (without colons) to a 64-bit unsigned integer
#-------------------------------------------------------------------------------------------------------#
function ConvertMacToULong {
    param($mac)
    return [Convert]::ToUInt64($mac, 16)
}

#-------------------------------------------------------------------------------------------------------#
# Function to check device info file exists
#-------------------------------------------------------------------------------------------------------#
function CheckForDeviceFile {
    Write-Host "--------------------------------------------------------------------------------------------------------------------------"
    Write-Host "Checking for device data file: $script:deviceFile..."
    if (Test-Path $script:deviceFile) {
        Write-Host "Device data file found."
        return $true
    } else {
        $msg = "Device data file not found. Please run FindBTDeviceInfo.ps1 first to create it."
        Write-Error $msg
        LogMessage $msg
        Start-Sleep 10
        throw $msg # Terminate script if no device file is found
    }
}

#-------------------------------------------------------------------------------------------------------#
# Function to import info from device file
#-------------------------------------------------------------------------------------------------------#
function ImportFromCSV {
    if (CheckForDeviceFile -eq $true) {
        Write-Host "Importing device information from $script:deviceFile..."
        $script:deviceData = Import-Csv -Path $script:deviceFile
        if (-not $script:deviceData -or $script:deviceData.Count -eq 0) {
            $msg = "No device data found in $script:deviceFile or the file is empty. Please run FindBTDeviceInfo.ps1."
            Write-Error $msg
            LogMessage $msg
            Start-Sleep 10
            throw $msg # Terminate script if device file is empty
        } else {
            Write-Host ""
            Write-Host "Device information loaded: Name='$($script:deviceData[0].DeviceName)', MAC='$($script:deviceData[0].MACAddress)'."
        }
    }
}

#-------------------------------------------------------------------------------------------------------#
# Function to check if device info is already loaded
#-------------------------------------------------------------------------------------------------------#

function DeviceInfoLoaded {
    if (-not $script:deviceData -or $script:deviceData.Count -eq 0) {
        
        #Run the ImportFromCSV function to load device information from the CSV file
        ImportFromCSV

        #------------------------------#
        # Prepare device info structure
        #------------------------------#
        # Convert the MAC string to a 64-bit address as required by the Bluetooth API
        $btAddr = ConvertMacToULong $script:deviceData[0].MACAddress

        # Instantiate the Bluetooth device structure and populate its fields 
        $script:info = New-Object "BLUETOOTH_DEVICE_INFO"
        $script:info.dwSize = 560  # Struct size in bytes
        $script:info.Address = $btAddr
        $script:info.fConnected = $false # Connection status isn't required for toggling service here 
        return $true 

    } else {
        return $true
    }
}

#-------------------------------------------------------------------------------------------------------#
# Function to check device status
#-------------------------------------------------------------------------------------------------------#
function DeviceStatus {
    if (DeviceInfoLoaded -eq $true) {
        $nameToCheck = $script:deviceData[0].DeviceName
        $statusOK = "OK"
        $foundAny = $false

        # Get all audio endpoint devices (connected or not)
        $devices = Get-PnpDevice -Class AudioEndpoint
        Write-Host "--------------------------------------------------------------------------------------------------------------------------"

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
}

#-------------------------------------------------------------------------------------------------------#
# Main Execution: Toggle A2DP Service
#-------------------------------------------------------------------------------------------------------#

$firstCheck = $true

while ($true) {
    $connected = DeviceStatus
    $friendlyName = $script:deviceData[0].DeviceName
    $deviceMAC = $script:deviceData[0].MACAddress

    if ($connected -eq $false) {
        if ($script:firstCheck -eq $true) {
            $script:firstCheck = $false
            Write-Host "Device has been disconnected!"
            Write-Host "Waiting $script:firstCheckInterval seconds before reconnecting."
            Start-Sleep $script:firstCheckInterval
        } else {
            try {
                Write-Host "Toggling A2DP service for $friendlyName (MAC: $deviceMAC)"
                # First, disable the A2DP service for the Bluetooth device
                [void][BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$script:a2dpGuid, 0)

                # Then, re-enable the A2DP service
                [void][BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$script:a2dpGuid, 1)
            }
            catch {
                # Capture and log any errors that occur during the toggle process
                LogMessage "Error occurred: $_"
            }
         }  
    } else {
        if ($script:firstCheck -eq $false) {
            $script:firstCheck = $true
        } 
    }
    Start-Sleep $script:interval
}
