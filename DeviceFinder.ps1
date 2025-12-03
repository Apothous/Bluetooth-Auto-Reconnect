function Get-AudioEndpointDeviceStatus {
    try {
        $devices = Get-PnpDevice -Class AudioEndpoint -ErrorAction Stop
    } catch {
        Write-Error "Failed to enumerate audio endpoints."
    }
    foreach ($device in $devices) {
        Write-Host "----------------------------------------"
        Write-Host "Device: $($device.FriendlyName)"
        Write-Host "Status: $($device.Status)"
    }
}
Get-AudioEndpointDeviceStatus
Write-Host "----------------------------------------"