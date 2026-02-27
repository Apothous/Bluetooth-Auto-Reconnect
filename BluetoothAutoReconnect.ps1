#-------------------------------------------------------------------------------------------------------#
# Script 2: Toggle Audio Service to connect PC to the Bluetooth Device
# Description:
#     This script disables and re-enables the A2DP (audio) service for a specific Bluetooth device.
#-------------------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------------------#
# User Configuration
#-------------------------------------------------------------------------------------------------------#

# Set the main loop interval
$script:toggleInterval = 30

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

$script:firstCheck = $true
$script:deviceInfoLoaded = $false

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
# Function to log current user/session for diagnostics
#-------------------------------------------------------------------------------------------------------#
function GetCurrentUser {
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    } catch {
        $currentIdentity = $env:USERNAME
    }
    #LogMessage "Running DeviceStatus as user '$currentIdentity'."
    Write-Host "Running DeviceStatus as user '$currentIdentity'."
}

#-------------------------------------------------------------------------------------------------------#
# Function to check device status
#-------------------------------------------------------------------------------------------------------#
function DeviceStatus {
    # Load device configuration if not already loaded
    if (-not $script:deviceInfoLoaded) {
        LogMessage "Attempting to load device information from $script:deviceFile..."
        if (-not (Test-Path $script:deviceFile)) {
            $msg = "Device data file not found. Please run SetupScript.ps1 first to create it."
            Write-Error $msg
            LogMessage $msg
            throw $msg # Terminate script if config file is missing
        }

        $csvData = Import-Csv -Path $script:deviceFile
        if (-not $csvData -or $csvData.Count -eq 0) {
            $msg = "No device data found in $script:deviceFile or the file is empty. Please run SetupScript.ps1."
            Write-Error $msg
            LogMessage $msg
            throw $msg # Terminate script if config file is empty
        }

        $script:friendlyName = $csvData[0].DeviceName
        $script:deviceMAC = $csvData[0].MACAddress
        
        #------------------------------#
        # Prepare device info structure
        #------------------------------#
        # Convert the MAC string to a 64-bit address as required by the Bluetooth API
        $btAddr = ConvertMacToULong $script:deviceMAC

        # Instantiate the Bluetooth device structure and populate its fields
        $script:info = New-Object "BLUETOOTH_DEVICE_INFO"
        $script:info.dwSize = 560  # Struct size in bytes
        $script:info.Address = $btAddr
        $script:info.fConnected = $false # Connection status isn't required for toggling service here

        $script:deviceInfoLoaded = $true
        LogMessage "Device information loaded: Name='$($script:friendlyName)', MAC='$($script:deviceMAC)'."
        Write-Host "-------------------------------------------------------------------------"
        Write-Host "Device information loaded: Name='$($script:friendlyName)', MAC='$($script:deviceMAC)'."
    }

    Write-Host "-------------------------------------------------------------------------"

    $nameToCheck = $script:friendlyName # Use the loaded friendly name
    $statusOK = "OK"
    $foundAny = $false
    $foundConnected = $false
    
    GetCurrentUser

    # Get all audio endpoint devices (connected or not) and guard against failures
    try {
        $devices = Get-PnpDevice -Class AudioEndpoint -ErrorAction Stop
    } catch {
        LogMessage "Get-PnpDevice failed: $($_ | Out-String)"
        Write-Error "Failed to enumerate audio endpoints. See log for details."
    }

    # Aggregate matching endpoints first. The device should be considered connected if ANY matching endpoint reports Status='OK'.
    $matchedDevices = @()
    foreach ($device in $devices) {
        $friendly = $device.FriendlyName
        $status = $device.Status

        $isMatch = $false
        if ($friendly) {
            if ($friendly -match "\(([^)]+)\)") {
                $nameInParens = $matches[1] -replace '\s+(Stereo|Hands-Free|Headset|AG Audio)$', ''
                if ($nameInParens -eq $nameToCheck) { $isMatch = $true }
            }
            # Fallback substring match if parenthesized form isn't present
            if (-not $isMatch -and ($friendly -like "*$nameToCheck*")) { $isMatch = $true }
        }

        if ($isMatch) {
            $matchedDevices += [PSCustomObject]@{ FriendlyName = $friendly; Status = $status }
        }
    }

    if ($matchedDevices.Count -gt 0) {
        $foundAny = $true
        $connectedCount = (@($matchedDevices | Where-Object { $_.Status -eq $statusOK })).Count
        if ($connectedCount -gt 0) {
            $foundConnected = $true
            Write-Host "Found $($matchedDevices.Count) matching '$nameToCheck' endpoint(s); $connectedCount connected."
        } else {
            Write-Host "Found $($matchedDevices.Count) matching '$nameToCheck' endpoint(s) but none are connected."
        }
    }

    if ($foundAny -eq $false) {
        Write-Host "Found no matching '$nameToCheck' endpoint(s); skipping toggle until device is present."
    }

    return $foundConnected
}

#-------------------------------------------------------------------------------------------------------#
# Function to toggle A2DP service
#-------------------------------------------------------------------------------------------------------#
function ToggleA2DPService {
    try {
        Write-Host "Toggling A2DP service for $script:friendlyName (MAC: $script:deviceMAC)"
        # First, disable the A2DP service for the Bluetooth device
        [void][BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$script:a2dpGuid, 0)
        # Then, re-enable the A2DP service
        [void][BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$script:a2dpGuid, 1)
        Write-Host "A2DP service toggled successfully."
    } catch {
        # Capture and log any errors that occur during the toggle process
        LogMessage "Error occurred while toggling A2DP service: $($_ | Out-String)"
        Write-Error "An error occurred while toggling A2DP service. See log for details."
    }
}

#-------------------------------------------------------------------------------------------------------#
# Event-Driven Bluetooth Reconnection Logic
#-------------------------------------------------------------------------------------------------------#

Register-WmiEvent -Query $Query -SourceIdentifier "BTDeviceEvent" -Action {
    try {
        Write-Host "[Event Triggered] Checking device connection status..."

        if (-not (DeviceStatus)) {
            Write-Host "$($script:friendlyName) is disconnected. Waiting $($script:firstCheckInterval) seconds..."
            Start-Sleep $script:firstCheckInterval
            Write-Host "Starting A2DP service toggle loop..."

            while (-not (DeviceStatus)) {
                ToggleA2DPService
                Start-Sleep $script:toggleInterval
            }

            Write-Host "$($script:friendlyName) successfully reconnected."
        } else {
            Write-Host "$($script:friendlyName) is connected. Waiting for next event."
        }

    } catch {
        LogMessage "Unexpected error in EventHandler: $($_ | Out-String)"
        Write-Error "Unexpected error in EventHandler: $($_.Exception.Message)"
    }
}

#-------------------------------------------------------------------------------------------------------#
# Main execution logic
#-------------------------------------------------------------------------------------------------------#

GetCurrentUser

# Start the event-driven reconnect function
EventHandler

# Wait indefinitely for events to trigger
Wait-Event -SourceIdentifier "BTDeviceEvent"
