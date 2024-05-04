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

Setting the Alert Time (voice):
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

FR:

Description

Ce script VBS (Visual Basic Script) permet à l'utilisateur de définir une alerte programmée à une heure spécifique. Lorsque l'heure approche, une notification est affichée pour rappeler l'utilisateur. Le programme vérifie également si plusieurs instances du même script sont en cours d'exécution et, dans ce cas, ferme les instances supplémentaires.


Prérequis

Windows OS avec support pour exécuter des scripts VBS.
Accès à l'éditeur de script ou à l'invite de commande pour exécuter le fichier.
Installation et Démarrage
Copiez le code dans un fichier texte.
Enregistrez le fichier avec l'extension .vbs, par exemple Fab_Alerte.vbs.
Double-cliquez sur le fichier pour lancer le script, ou exécutez-le via l'invite de commande avec wscript Fab_Alerte.vbs.


Utilisation

Définir l'Heure d'Alerte (vocale):
Au démarrage, une boîte de dialogue s'ouvre pour vous demander d'entrer l'heure de l'alerte au format HH:MM.
Un exemple basé sur l'heure actuelle plus 6 minutes est affiché pour aider à formuler l'heure correctement.
Si le format de l'heure est incorrect ou si l'heure est trop proche (moins de 5 minutes dans le futur), un message d'erreur sera affiché et vous serez invité à réessayer.
Notification d'Alerte :
Une fois une heure valide saisie, une notification préventive s'affichera 5 minutes avant l'heure d'alerte programmée.
Un rappel final s'affiche exactement à l'heure programmée.
Répétition :
Après l'alerte, le programme invite l'utilisateur à définir une nouvelle heure pour la prochaine alerte.


Fonctions Spéciales

CheckIfRunning : Vérifie si plusieurs instances du script sont en cours d'exécution. Si c'est le cas, les instances supplémentaires seront fermées pour éviter les doublons.
CustomSleep : Une fonction personnalisée pour mettre le script en pause tout en vérifiant périodiquement si le script doit être terminé prématurément.
Sortir : Ferme le script avec un message si plusieurs instances sont détectées.


Dépannage

Script ne démarre pas : Vérifiez si les scripts VBS sont autorisés sur votre système et si vous avez les droits d'administrateur nécessaires.
Erreurs de format d'heure : Assurez-vous de suivre le format HH:MM et que l'heure est au moins 5 minutes dans le futur.
