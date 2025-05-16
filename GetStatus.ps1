Get-PnpDevice -Class AudioEndpoint | ForEach-Object {
    Write-Host "FriendlyName: $($_.FriendlyName)"
    Write-Host "Status: $($_.Status)"
    Write-Host "InstanceId: $($_.InstanceId)"
    Write-Host "------------------------------------"
}
