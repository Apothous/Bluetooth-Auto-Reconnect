Steps for setting up Bluetooth Auto Reconnect Script.

------------------------
----- Setup Script -----
------------------------

1. Open File Explorer and navigate to the USB drive directory containing the Bluetooth Auto Reconnect folder. 

2. Copy the Bluetooth Auto Reconnect folder from USB drive and paste it into the C:\ directory.

3. Right click on the SetupScript.ps1 and select "Run With PowerShell" (If this option isn't available then search for PowerShell in the Windows Search bar and open it with administrator privileges, once prompted run the following command: C:\Bluetooth Auto Reconnect\SetupScript.ps1)
	- This script will search the PC and list all the paired Bluetooth devices as well as create two files in the Bluetooth Auto Reconnect folder (If they don't already exist).
		- Devices.csv [This file contains a list of the Names and Mac Address's for all paired Bluetooth devices]
		- LogFile.txt [This file contains a timestamped log of all the actions taken by both scripts]
	- Next it will prompt the user to choose a device from the list of paired Bluetooth devices and save the device info to the CSV file.
	- It will then prompt the user to create a scheduled task to run the Bluetooth Auto Reconnect script.

-------------------------------------------
----- Test that the script is working -----
-------------------------------------------
 
1. First make sure the Bluetooth device is turned on and available to connect (in pairing/connection mode) by pressing and holding the pairing button on the Bluetooth device.

2. Pair the Bluetooth device with the internal PC using the Windows Action Center and make sure it fully connects. (Optional - test for audio)

3. Disconnect the Bluetooth device from the internal PC.

4. Find the newly created Bluetooth Auto Reconnect task in the Task Scheduler and run the task. Wait (the specified time interval) and make sure that the Bluetooth device has reconnected to the internal PC. 

5. Press and hold the pairing button on the Bluetooth device again to disconnect it from the internal PC and then immediately pair the device with a secondary PC, Phone , or Tablet and verifying the connection. (Optional - test for audio)

6. Disconnect the device from the secondary PC, Phone, or Tablet (this can be done from the secondary device or with the pairing button on the Bluetooth device) and wait (the specified time interval) to make sure that the Bluetooth device has reconnected to the internal PC. (Optional - test for audio)

7. (Optional) Repeat steps 3-6 at least one more time to confirm consistent operation.
