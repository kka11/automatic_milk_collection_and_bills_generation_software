# MDBCS-Software
Database conversion and Billing Automation Software

This software is developed by Harshit Singh and Ankit Kumar students of IIT Jodhpur for ICICI RSETI.

First Watch the Demo Video

Directions for use:

1. Double click the Conversion Application installer.exe
2. And Install the Software.
3. Copy the "Database" folder in the "D:\" drive or make a directory containing the structure same as "Database" folder.   
   If you are changing the directory of the "Database" folder kindly update the same(directory) in the config.cfg file in the installation folder.
4. Go to the "Conversion Application" folder (or where you have installed the Software) 
5. Double click the "MDBCS.exe" to run the Milk Database Conversion Software.
6. Choose the Mode of Sheet generation from "By Date" or "By Payment Cycle"
7. For "By Date" option - Enter the center code, Starting Date and End Date 
	the proper format for date entry is - yyyy-mm-dd 
	for example - 2017-08-31 
   and click "Convert".
8. For "By Payment Cycle" option - Enter the center code and click "By Payment".
9. Go to "Database\<center code>\" to find the "Final Sheet" generated.
10. You can also switch between the two mode at any point of time.

The "By Payment Cycle" option generates the sheet in the following manner:
	for date : 1-10 of the current month of the PC's clock if the current date is 11-20 of the current month 
	for date : 11-20 of the current month of the PC's clock if the current date is 21-31 of the current month
	for date : 21-31 of the current month of the PC's clock if the current date is 1-10 of the current month 
  
