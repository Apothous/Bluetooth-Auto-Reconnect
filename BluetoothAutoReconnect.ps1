#-------------------------------------------------------------------------------------------------------#
# Script 2: Toggle Audio Service to connect PC to the Bluetooth Device
# Description:
#     This script disables and re-enables the A2DP (audio) service for a specific Bluetooth device.
#-------------------------------------------------------------------------------------------------------#

#-------------------------------------------------------------------------------------------------------#
# Interval Settings
#-------------------------------------------------------------------------------------------------------#

# Set the main loop interval
$script:toggleInterval = 60

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

#Check for first instance of disconnected status before attempting to toggle A2DP service
$script:firstCheck = $true

#-------------------------------------------------------------------------------------------------------#
# Type Definition for Bluetooth Device Structure and Service Manager
#-------------------------------------------------------------------------------------------------------#

# Define the necessary C# structures and import Bluetooth function from Windows API (if not already loaded)
if (-not ([System.Management.Automation.PSTypeName]'BLUETOOTH_DEVICE_INFO').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential)]
    public struct SYSTEMTIME {
        public ushort wYear;
        public ushort wMonth;
        public ushort wDayOfWeek;
        public ushort wDay;
        public ushort wHour;
        public ushort wMinute;
        public ushort wSecond;
        public ushort wMilliseconds;
    }

    // This structure represents a Bluetooth device's key info
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct BLUETOOTH_DEVICE_INFO {
        public uint dwSize;           // Size of this structure in bytes (required by API)
        public ulong Address;         // Bluetooth address (converted from MAC string)
        public uint ulClassofDevice; // Device class (not used here, but required for struct)
        public int fConnected;       // BOOL (4 bytes) - Whether the device is currently connected
        public int fRemembered;      // BOOL (4 bytes) - Whether the device is remembered (paired)
        public int fAuthenticated;   // BOOL (4 bytes) - Whether the device is authenticated (paired)
        public SYSTEMTIME stLastSeen; // Last time the device was seen (not used here)
        public SYSTEMTIME stLastUsed; // Last time the device was used (not used here)
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 248)]
        public string szName;        // Device name (not used here, but required for struct)
    }

    // Provides access to the BluetoothSetServiceState WinAPI function
    public class BtServiceManager {
        [DllImport("bthprops.cpl", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern uint BluetoothSetServiceState(
            IntPtr hRadio,                          // Not used (set to zero)
            ref BLUETOOTH_DEVICE_INFO deviceInfo,   // Info struct for target device
            ref Guid guidService,                   // Service GUID (e.g., A2DP)
            uint dwServiceFlags                     // 0 = disable, 1 = enable
        );

        [DllImport("bthprops.cpl", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern uint BluetoothGetDeviceInfo(
            IntPtr hRadio,                          // Not used (set to zero)
            ref BLUETOOTH_DEVICE_INFO pbtdi    // Info struct for target device
        );
    }
"@ -Language CSharp
}

#-------------------------------------------------------------------------------------------------------#
# Function to log a timestamped message to the log file
#-------------------------------------------------------------------------------------------------------#
function LogMessage {
    param([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -Append $script:logFile
}

#-------------------------------------------------------------------------------------------------------#
# Function to load device information from the CSV file
#-------------------------------------------------------------------------------------------------------#
function LoadDeviceInfo {
    LogMessage "Attempting to load device information from $script:deviceFile..."
    
    # Load device configuration
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

    Write-Host "-------------------------------------------------------------------------"
    Write-Host "Device information loaded: Name='$($script:friendlyName)', MAC='$($script:deviceMAC)'."
}

#-------------------------------------------------------------------------------------------------------#
# Function to check device status
#-------------------------------------------------------------------------------------------------------#
function DeviceStatus {
    
    Write-Host "-------------------------------------------------------------------------"
    
    #-------------------------------#
    # Prepare device info structure #
    #-------------------------------#

    # Convert the MAC string to a 64-bit address as required by the Bluetooth API
    try {
        $btAddr = [Convert]::ToUInt64($script:deviceMAC, 16)
    } catch {
        $msg = "Invalid MAC address format in CSV: '$script:deviceMAC'. Please re-run SetupScript.ps1."
        LogMessage $msg
        Write-Error $msg
        throw $msg
    }

    # Instantiate the Bluetooth device structure and populate its fields
    $script:info = New-Object "BLUETOOTH_DEVICE_INFO"
    $script:info.dwSize = 560  # Struct size in bytes
    $script:info.Address = $btAddr

    #----------------------------------------------#
    # Use Bluetooth API to check connection status #
    #----------------------------------------------#

    Write-Host "Checking Bluetooth connection status for $script:friendlyName..."
    try {
        $deviceStatus = [BtServiceManager]::BluetoothGetDeviceInfo([IntPtr]::Zero, [ref]$script:info)
    } catch {
        LogMessage "Error calling BluetoothGetDeviceInfo: $($_.Exception.Message)"
        Write-Host "Error calling Bluetooth API. Assuming $script:friendlyName is disconnected."
        return $false
    }

    if ($deviceStatus -eq 0) {
        # Safety Switch: Check if the device is still paired (Remembered)
        if ($script:info.fRemembered -eq 0) {
            $msg = "FATAL: Device '$script:friendlyName' is no longer paired (Remembered). Stopping script."
            LogMessage $msg
            Write-Error $msg
            exit # Terminate the script entirely to prevent log spam
        }

        if ($script:info.fConnected -ne 0) {
            $isConnected = $true
            Write-Host "$script:friendlyName Status: Connected"
        } else {
            $isConnected = $false
            Write-Host "$script:friendlyName Status: Disconnected"
        }
        return $isConnected
    } else {
        Write-Host "BluetoothGetDeviceInfo returned error code: $deviceStatus."
        Write-Host "Assuming $script:friendlyName is disconnected."
        return $false
    }
}

#-------------------------------------------------------------------------------------------------------#
# Function to toggle A2DP service
#-------------------------------------------------------------------------------------------------------#
function ToggleA2DPService {
    try {
        Write-Host "Toggling A2DP service for $script:friendlyName..."

        # First, disable the A2DP service for the Bluetooth device
        $disableResult = [BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$script:a2dpGuid, 0)
        if ($disableResult -ne 0) { Write-Warning "Disable A2DP failed with error code: $disableResult" }
        
        Start-Sleep 2 # Short pause to allow the system to process the service state change
        
        # Then, re-enable the A2DP service
        $enableResult = [BtServiceManager]::BluetoothSetServiceState([IntPtr]::Zero, [ref]$script:info, [ref]$script:a2dpGuid, 1)
        if ($enableResult -ne 0) { Write-Warning "Enable A2DP failed with error code: $enableResult" }
        
        Write-Host "A2DP service toggled successfully."
    } catch {
        # Capture and log any errors that occur during the toggle process
        LogMessage "Error occurred while toggling A2DP service: $($_ | Out-String)"
        Write-Error "An error occurred while toggling A2DP service. See log for details."
    }
}

#-------------------------------------------------------------------------------------------------------#
# Main execution logic
#-------------------------------------------------------------------------------------------------------#
LoadDeviceInfo
while ($true) {
    try {
        $connected = DeviceStatus
        if ($connected -eq $true -and $script:firstCheck -eq $false) {
            $script:firstCheck = $true
        } elseif ($connected -eq $false -and $script:firstCheck -eq $true) {
            $script:firstCheck = $false
            Write-Host "Waiting $script:firstCheckInterval seconds before attempting to reconnect."
            Start-Sleep $script:firstCheckInterval
        } elseif ($connected -eq $false -and $script:firstCheck -eq $false) {
            Write-Host "Device is still disconnected. Attempting to toggle A2DP service..."
            ToggleA2DPService
        }
    } catch {
        LogMessage "Unexpected error in main loop: $($_.Exception.Message)"
        Write-Host "An unexpected error occurred. Retrying in $script:toggleInterval seconds."
    }

    Start-Sleep $script:toggleInterval
}