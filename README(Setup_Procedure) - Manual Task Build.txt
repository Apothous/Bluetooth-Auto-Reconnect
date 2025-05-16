Steps for setting up Bluetooth Auto Reconnect Script.

------------------------
----- Setup Script -----
------------------------

1. Open File Explorer and navigate to the USB drive directory containing the Bluetooth Auto Reconnect folder. 

2. Copy the Bluetooth Auto Reconnect folder from USB drive and paste it into the C:\ directory.

3. Right click on the FindBTDeviceInfo.ps1 and select "Run With PowerShell" (If this option isn't available then search for PowerShell in the Windows Search bar and open it, once prompted run the following command: C:\Bluetooth Auto Reconnect\FindBTDeviceInfo.ps1
	- This script will search the PC and list all the paired Bluetooth devices as well as create two files in the Bluetooth Auto Reconnect folder (If they don't already exist).
		- Devices.csv [This file contains a list of the Names and Mac Address's for all paired Bluetooth devices]
		- LogFile.txt [This file contains a timestamped log of all the actions taken by both scripts]

4. Open the Devices.csv file to retrieve the Device Name and MAC Address for the desired device.

5. After obtaining the desired device information, open the BluetoothAutoReconnect.ps1 script in a text editor and change the "friendlyName" and "deviceMAC" variables located at the top of the script to the desired device (Letters in MAC Address should all be capitalized and there should be no separating Colons).

--------------------------------
----- Setup Scheduled Task -----
--------------------------------

1. Open Windows Search, search for Task Scheduler (taskschd.msc) and open it.

2. Create a new task 
	- Click "Action"
	- Create "Task"

3. In the Create Task Window 
	- Name the task Bluetooth Auto Reconnect
	- Check the "Run with Highest Privileges" checkbox 
	- Check the "Hidden" box (to prevent popup window)
	- Click on the "Configure for" dropdown menu and select "Windows 10"

4. Copy and Paste the following description into the description box:
This task continuously runs a PowerShell script called Bluetooth Auto Reconnect every minute. The script enables and disables the A2DP service for a paired Bluetooth audio device specified using the devices MAC Address. By toggling the A2DP service the PC sends a signal to connect with the paired Bluetooth audio device if the device is within range and available to connect.

5. Navigate to the "Triggers" tab and click on the "New" button

6. In the "New Trigger" window 
	- Click on the "Begin the task" drop down menu and select "At startup" 
	- Uncheck the "Repeat task every" box
	- Click the "OK" button

7. Navigate to the "Actions" tab and click on the "New" button

8. In the "New Action" window 
	- Click on the "Action" drop down menu and select "Start a program" 
	- In the "Program/script" textbox type "powershell.exe" 
	- In the "Add arguments (optional)" text box copy and past the following: -WindowStyle Hidden -ExecutionPolicy Bypass -File "D:\Bluetooth\Bluetooth Auto Reconnect\BluetoothAutoReconnect.ps1"
	- Click the "OK" button

9. Navigate to the "Conditions" tab and unselect the "Start the task only if the computer is on AC power" checkbox

10. Navigate to the "Settings" tab and make sure the settings are as follows
	- [Checked] "Allow task to be run on demand"
	- [Checked] "Run task as soon as possible after a scheduled restart is missed"
	- [Checked] "If the task fails, restart every" (make sure settings are 1 minute and 3 times)
	- [Unchecked] "Stop the task if it runs longer than"
	- [Unchecked] "If the running task does not end when requested, force it to stop"
	- [Unchecked] "If the task is not scheduled to run again, delete it after"
	- In the "If the task is already running, then the following rule applies" dropdown menu select "Do not start a new instance"

11. Click the "OK" button to save the scheduled task

---------------------------------------
----- Test that script is working -----
---------------------------------------
 
1. First make sure the Bluetooth device is turned on and available to connect (in pairing/connection mode) by pressing and holding the pairing button on the Bluetooth device.

2. Pair the Bluetooth device with the internal PC using the Windows Action Center and make sure it fully connects. (Optional - test for audio)

3. Disconnect the Bluetooth device from the internal PC.

4. Find the newly created Bluetooth Auto Reconnect task in the Task Scheduler and run the task. Wait (the specified time interval) and make sure that the Bluetooth device has reconnected to the internal PC. 

5. Press and hold the pairing button on the Bluetooth device again to disconnect it from the internal PC and then immediately pair the device with a secondary PC, Phone , or Tablet and verifying the connection. (Optional - test for audio)

6. Disconnect the device from the secondary PC, Phone, or Tablet (this can be done from the secondary device or with the pairing button on the Bluetooth device) and wait (the specified time interval) to make sure that the Bluetooth device has reconnected to the internal PC. (Optional - test for audio)

7. (Optional) Repeat steps 3-6 at least one more time to confirm consistent operation.








