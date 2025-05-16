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

2. In the left panel look for and click on "Import Task".

3. Navigate to the C:\Bluetooth Auto Reconnect directory.

4. Select the Bluetooth Auto Reconnect.xml file and click "Open".

5. Click the "OK" button in the "Create Task" pop up window.

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
