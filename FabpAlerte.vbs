Option Explicit

Dim objWMIService, colProcesses
Dim strProcessName, strComputer
Dim currentDate

currentDate = FormatDateTime(Now, vbShortDate) ' Formate la date actuelle en format court.
strProcessName = "wscript.exe"  ' ou "cscript.exe" selon l'interpréteur utilisé
strComputer = "."

' Créer une connexion au service WMI
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

' Appel de la fonction principale
Call Main

Sub Main
    Dim alertTime, preAlertTime, userTimeInput, isTimeValid, exampleTime, minutesRemaining
	If CheckIfRunning() > 1 Then
               Sortir()
            End If
    isTimeValid = False
    exampleTime = FormatDateTime(DateAdd("n", 6, Now), vbShortTime)  ' Calculer l'exemple de l'heure actuelle plus 6 minutes

    Do
        userTimeInput = InputBox("Veuillez entrer l'heure de l'alerte (format HH:MM, par exemple " & exampleTime & ")", "Fab_Alerte")
        If userTimeInput = "" Then ' Si l'utilisateur annule la saisie
            MsgBox "Aucune heure entrée, le programme va s'arrêter.", vbOKOnly + vbExclamation, "Fab_Alerte Programme Terminé"
            Exit Sub
        End If

        On Error Resume Next
        alertTime = CDate(currentDate & " " & userTimeInput & ":00") ' Ajouter les secondes pour compléter le format HH:MM:SS
        MsgBox "  - Votre saisie : " & alertTime &"  - Maintenant : " & Now,vbOKOnly
		If Err.Number <> 0 Then
            MsgBox "Format d'heure invalide, veuillez réessayer.", vbOKOnly + vbCritical, "Erreur"
            Err.Clear
        ElseIf Now > alertTime Or DateDiff("n", Now, alertTime) < 5 Then
            MsgBox "L'heure entrée doit être au moins 5 minutes plus tard que l'heure actuelle.", vbOKOnly + vbExclamation, "Heure invalide"
        Else
            isTimeValid = True
        End If
        On Error GoTo 0
    Loop Until isTimeValid

    preAlertTime = DateAdd("n", -5, alertTime) ' Configurer l'alerte pour 5 minutes avant l'heure entrée

    Do While Now < alertTime
        If Now >= preAlertTime Then
         

            minutesRemaining = DateDiff("n", Now, alertTime)
            MsgBox "Ceci est un rappel! Il reste maintenant " & minutesRemaining & " minutes avant l'heure de votre alerte.", vbOKOnly + vbInformation, " Fab_Alerte Notification"
            CustomSleep(150) ' Attendre 60 secondes avec vérification
        Else
            CustomSleep(30) ' Vérifier toutes les 30 secondes
        End If
    Loop

    MsgBox "L'heure de l'alerte est passée. Veuillez entrer une nouvelle heure pour la prochaine alerte.", vbOKOnly + vbInformation, "Fab_Alerte Alerte Terminée"
    Call Main ' Demander une nouvelle heure pour la prochaine alerte
End Sub

Function CheckIfRunning()
    Dim item, instanceCount
    instanceCount = 0
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'")
    For Each item In colProcesses
        If InStr(item.CommandLine, WScript.ScriptName) > 0 Then
            instanceCount = instanceCount + 1
        End If
    Next
    CheckIfRunning = instanceCount
End Function

Sub CustomSleep(seconds)
    Dim start, elapsed
    start = Now
    Do While DateDiff("s", start, Now) < seconds
        If CheckIfRunning() > 1 Then
             Sortir()
        End If
        WScript.Sleep(1000) ' Sleep for 1 second
    Loop
End Sub

Sub Sortir()
    
              WScript.Echo "Une autre instance du programme est détectée. Ce programme va s'arrêter après 10 secondes."
                WScript.Sleep(10000)  ' Attendre 10 secondes avec vérification
                WScript.Quit
 
End Sub
