# HP-Dock-Serial-Number-Grabber
Powershell script to grab the serial number, firmware version and other information regarding a connected HP G5 USB-C Docking Station

This script was written to help aide in inventorying HP G5 USB-C Docking stations by collecting the relavent information for each device. This script does require that the laptop running the script be plugged into each docking station. 

## To Use
To use this script, you must have the HPDockWMIProvider software installed on the host machine. This is required to get the information from the WMI calls.
Once HPDockWMIProvider is installed,, running the script will loop until you end the script. Each loop will ask to confirm that the device has connected to the next docking station, the room the docking station is in, and the asset tag number. 

The script will know you are done collecting data when you select "No" to the prompt "Connected to Dock?" Once this has been selected, the data collected will be shown on the screen, and you will be asked if you want to save the data.
