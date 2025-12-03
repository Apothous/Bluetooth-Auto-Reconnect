# Bluetooth-Auto-Reconnect — Copilot instructions

These notes help an AI coding agent be productive quickly in this small PowerShell project.

## Big picture
- Two main scripts:
  - `SetupScript.ps1` — enumerates paired Bluetooth devices, prompts the user to pick one, writes `BTDevice.csv`, and optionally registers a Scheduled Task.
  - `BluetoothAutoReconnect.ps1` — long-running watchdog that reads `BTDevice.csv`, checks connection status and toggles the A2DP service via the Windows Bluetooth API to force reconnects.
- Data flow: `SetupScript.ps1` -> `BTDevice.csv` (DeviceName, MACAddress) -> `BluetoothAutoReconnect.ps1` (reads CSV, uses MAC to build `BLUETOOTH_DEVICE_INFO`).

## Key files to inspect
- `SetupScript.ps1` — registry scanning (`HKLM:\SYSTEM\CurrentControlSet\Services\BTHPORT\Parameters\Devices`), CSV creation, scheduled task registration code.
- `BluetoothAutoReconnect.ps1` — runtime loop, `Add-Type` C# definition for `BLUETOOTH_DEVICE_INFO` and `BluetoothSetServiceState`, `Get-PnpDevice -Class AudioEndpoint` usage.
- `README(Setup_Procedure).txt` and `LogicExplanation.txt` — user-facing setup and rationale; good for examples and expected user flows.

## Project-specific conventions and implementation details
- Files are referenced relative to the script folder via `$PSScriptRoot`. Use that when adding or modifying file paths.
- `BTDevice.csv` columns: `DeviceName`, `MACAddress` (MAC stored uppercase without separators). Do not change the CSV shape.
- The A2DP GUID is hard-coded as `0000110b-0000-1000-8000-00805f9b34fb` and used with `BluetoothSetServiceState` (from `bthprops.cpl`).
- `BLUETOOTH_DEVICE_INFO` is added via `Add-Type` C# on script load; the script sets `.dwSize = 560` — preserve this unless you validate the native layout.
- Windows version branch: `IsWindows11` checks build >= 22000 and alters how `FriendlyName` is parsed from `Get-PnpDevice` results.
- Long-running loop pattern: `while ($true) { ... Start-Sleep $script:interval }` with two configurable globals: `$script:interval` (default 30) and `$script:firstCheckInterval` (default 60).

## Developer workflows (how to run & test)
- Run the setup script (requires elevation):
  - In an elevated PowerShell prompt: `PowerShell -NoProfile -ExecutionPolicy Bypass -File "C:\Bluetooth Auto Reconnect\SetupScript.ps1"`
  - Follow prompts to write `BTDevice.csv` and optionally register the Scheduled Task.
- Run the watchdog directly (for debugging):
  - Elevated PowerShell: `PowerShell -NoProfile -ExecutionPolicy Bypass -File "C:\Bluetooth Auto Reconnect\BluetoothAutoReconnect.ps1"`
  - Stop the loop with Ctrl+C while debugging.
- Scheduled task: task name is `Bluetooth Auto Reconnect` and is registered to run as `NT AUTHORITY\SYSTEM` with triggers `AtStartup` and `AtLogOn`. The action runs `powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "<script>"`.

## Debugging tips and safe editing guidance
- When changing native interop (`Add-Type`) or `dwSize`, validate with small tests first — incorrect struct layout will cause silent native failures.
- For quicker iteration, import the functions in an interactive session (dot-source the script) and call `ToggleA2DPService` or `DeviceStatus` manually.
- Logs are appended to `LogFile.txt` in the script folder via `LogMessage`. Use the log for runtime errors and to confirm scheduled task behavior.
- `Get-PnpDevice -Class AudioEndpoint` is central to detecting connection state; test on both Windows 10 and Windows 11 VMs if behavior differs.

## Things NOT to change without validation
- The Bluetooth A2DP GUID value and the callsite to `BluetoothSetServiceState`.
- CSV schema and the way MACs are converted with `ConvertMacToULong` (expects hex string without separators).
- Scheduled task principal and triggers unless the change is explicitly desired.

## Example edits an agent might perform
- Add a `--dry-run` flag to `BluetoothAutoReconnect.ps1` that logs intended actions without calling `BluetoothSetServiceState`.
- Parameterize `$script:interval` and `$script:firstCheckInterval` to accept environment variables or CLI args.
- Improve robustness: add explicit checks around `Get-PnpDevice` and surface clearer log messages for parsing failures.

If anything above is unclear or you want examples (patches/tests) for any suggested edits, tell me which area to expand. 
