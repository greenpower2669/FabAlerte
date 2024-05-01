# FabAlerte
popup reminder program

Description
This VBS (Visual Basic Script) script allows users to set a scheduled alert for a specific time. When the time approaches, a notification is displayed to remind the user. The program also checks if multiple instances of the same script are running and, if so, closes the additional instances.


Requirements

Windows OS with support for running VBS scripts.
Access to a script editor or command prompt to execute the file.
Installation and Startup
Copy the code into a text file.
Save the file with the .vbs extension, for example, Fab_Alerte.vbs.
Double-click the file to launch the script or run it via command prompt with wscript Fab_Alerte.vbs.


Usage

Setting the Alert Time:
Upon startup, a dialog box will prompt you to enter the alert time in HH:MM format.
An example based on the current time plus 6 minutes is displayed to help correctly formulate the time.
If the time format is incorrect or the time is too close (less than 5 minutes in the future), an error message will be displayed and you will be prompted to retry.
Alert Notification:
Once a valid time is entered, a preemptive notification will be displayed 5 minutes before the scheduled alert time.

A final reminder appears exactly at the scheduled time.


Repetition:

After the alert, the program prompts the user to set a new time for the next alert.


Special Functions

CheckIfRunning: Checks if multiple instances of the script are running. If so, additional instances will be closed to avoid duplicates.
CustomSleep: A custom function to pause the script while periodically checking if the script should be terminated prematurely.
Sortir: Closes the script with a message if multiple instances are detected.


Troubleshooting

Script doesn't start: Ensure that VBS scripts are allowed on your system and that you have the necessary administrator rights.
Time format errors: Make sure to follow the HH:MM format and that the time is at least 5 minutes in the future.
This user manual will help you set up and effectively use the popup reminder program to remind you of important scheduled events.
