Option Explicit

Dim objWMIService, colProcesses
Dim strProcessName, strComputer
Dim currentDate
Dim message
Set message = CreateObject("SAPI.SpVoice")


currentDate = FormatDateTime(Now, vbShortDate) ' Formate la date actuelle en format court.
strProcessName = "wscript.exe"  ' ou "cscript.exe" selon l'interpréteur utilisé
strComputer = "."

' Créer une connexion au service WMI
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

' Appel de la fonction principale
Call Main

Sub Main
    Dim alertTime, preAlertTime, userTimeInput, isTimeValid, exampleTime, minutesRemaining,nowAlertTime
	If CheckIfRunning() > 1 Then
               Sortir()
            End If
    isTimeValid = False
    exampleTime = FormatDateTime(DateAdd("n", 6, Now), vbShortTime)  ' Calculer l'exemple de l'heure actuelle plus 6 minutes

    Do
		message.Speak "Bien, information FabAlert: Veuillez entrer l'heure de l'alerte"
        userTimeInput = InputBox("Veuillez entrer l'heure de l'alerte (format HH:MM, par exemple " & exampleTime & ")", "Fab_Alerte")
        If userTimeInput = "" Then ' Si l'utilisateur annule la saisie
			message.Speak "Bien, information FabAlert: Aucune heure entrée, le programme va s'arrêter."
            MsgBox "Aucune heure entrée, le programme va s'arrêter.", vbOKOnly + vbExclamation, "Fab_Alerte Programme Terminé"
            Exit Sub
        End If

        On Error Resume Next
        alertTime = CDate(currentDate & " " & userTimeInput & ":00") ' Ajouter les secondes pour compléter le format HH:MM:SS
		
        
		
		If Err.Number <> 0 Then
		message.Speak "Bien, information FabAlert: Format d'heure invalide, veuillez réessayer."
            MsgBox "Format d'heure invalide, veuillez réessayer.", vbOKOnly + vbCritical, "Erreur"
			
            Err.Clear
        ElseIf Now > alertTime Or DateDiff("n", Now, alertTime) < 5 Then
			message.Speak "Bien, information FabAlert: L'heure entrée doit être au moins 5 minutes plus tard que l'heure actuelle."
            MsgBox "L'heure entrée doit être au moins 5 minutes plus tard que l'heure actuelle.", vbOKOnly + vbExclamation, "Heure invalide"
			
        Else
            isTimeValid = True
			message.Speak " Bien, information FabAlert:  - Votre saisie : " & alertTime & " est valide!"
			MsgBox "  - Votre saisie : " & alertTime &"  - Maintenant : " & Now,vbOKOnly
        End If
        On Error GoTo 0
    Loop Until isTimeValid
	nowAlertTime = DateAdd("n", -1, alertTime) ' Configurer l'alerte pour 5 minutes avant l'heure entrée
    preAlertTime = DateAdd("n", -5, alertTime) ' Configurer l'alerte pour 5 minutes avant l'heure entrée

    Do While Now < nowAlertTime'alertTime
        If Now >= preAlertTime Then
         

            minutesRemaining = DateDiff("n", Now, alertTime)
			message.Speak "Bien, information FabAlert: Ceci est un rappel! Il reste maintenant " & minutesRemaining & "minutes avant l'heure de votre alerte."
            MsgBox "Ceci est un rappel! Il reste maintenant " & minutesRemaining & " minutes avant l'heure de votre alerte.", vbOKOnly + vbInformation, " Fab_Alerte Notification"
            CustomSleep(150) ' Attendre 60 secondes avec vérification
        Else
            CustomSleep(30) ' Vérifier toutes les 30 secondes
        End If
    Loop
		message.Speak "Bien, information FabAlert: L'heure de l'alerte est passée. Veuillez entrer une nouvelle heure pour la prochaine alerte"
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
				message.Speak "Bien, information FabAlert: Une autre instance du programme est détectée. Ce programme va s'arrêter après 10 secondes."
              WScript.Echo "Une autre instance du programme est détectée. Ce programme va s'arrêter après 10 secondes."
                WScript.Sleep(10000)  ' Attendre 10 secondes avec vérification
                WScript.Quit
 
End Sub
