# Define the path to the CSV file where Bluetooth device information was saved
$deviceFile = "C:\Bluetooth Auto Reconnect\Devices.csv"

function DeviceStatus {
    # Import the CSV and grab the DeviceName of the first entry
    $csvData = Import-Csv -Path $deviceFile

    if (-not $csvData -or $csvData.Count -eq 0) {
        Write-Host "No device data found in $deviceFile"
        return $false
    }

    $nameToCheck = $csvData[0].DeviceName  # Use the first device in the list (or loop if needed)
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
                    Write-Host "Device is connected."
                    return $true
                } else {
                    Write-Host "Device is disconnected"
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
