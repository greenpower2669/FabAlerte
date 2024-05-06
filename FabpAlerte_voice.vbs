Option Explicit

Dim objWMIService, colProcesses
Dim strProcessName, strComputer
Dim currentDate
Dim message
dim oShell
Set message = CreateObject("SAPI.SpVoice")
Set oShell = CreateObject("WScript.Shell")

currentDate = FormatDateTime(Now, vbShortDate) ' Formate la date actuelle en format court.
strProcessName = "wscript.exe"  ' ou "cscript.exe" selon l'interpr�teur utilis�
strComputer = "."

' Cr�er une connexion au service WMI
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
		message.Speak "Bien, annonce FabAlert : Veuillez entrer l'heure de l'alerte"
        userTimeInput = InputBox("Veuillez entrer l'heure de l'alerte (format HH:MM, par exemple " & exampleTime & ")", "Fab_Alerte")
        If userTimeInput = "" Then ' Si l'utilisateur annule la saisie
			message.Speak "Bien, annonce FabAlert : Aucune heure entr�e, le programme va s'arr�ter."
            'MsgBox "Aucune heure entr�e, le programme va s'arr�ter.", vbOKOnly + vbExclamation, "Fab_Alerte Programme Termin�"
			'convert in temporary shell because of voice alerte, that is enouth
			oShell.Popup "Aucune heure entr�e, le programme va s'arr�ter.", 5, "Fab_Alerte Programme Termin�",  4096 +16
			


            Exit Sub
        End If

        On Error Resume Next
        alertTime = CDate(currentDate & " " & userTimeInput & ":00") ' Ajouter les secondes pour compl�ter le format HH:MM:SS
		
        
		
		If Err.Number <> 0 Then
		message.Speak "Bien, annonce FabAlert : Format d'heure invalide, veuillez r�essayer."
            'MsgBox "Format d'heure invalide, veuillez r�essayer.", vbOKOnly + vbCritical, "Erreur"
			oShell.Popup "Format d'heure invalide, veuillez r�essayer." , 5, "Erreur",  4096 +16
			
            Err.Clear
        ElseIf Now > alertTime Or DateDiff("n", Now, alertTime) < 5 Then
			message.Speak "Bien, annonce FabAlert : L'heure entr�e doit �tre au moins 5 minutes plus tard que l'heure actuelle."
            'MsgBox "L'heure entr�e doit �tre au moins 5 minutes plus tard que l'heure actuelle.", vbOKOnly + vbExclamation, "Heure invalide"
			oShell.Popup "L'heure entr�e doit �tre au moins 5 minutes plus tard que l'heure actuelle." , 5, "Heure invalide",  4096 +16
		
			
			
			
        Else
            isTimeValid = True
			message.Speak " Bien, annonce FabAlert :  - Votre saisie : " & alertTime & " est valide!"
			'MsgBox "  - Votre saisie : " & alertTime &"  - Maintenant : " & Now,vbOKOnly
			oShell.Popup "  - Votre saisie : " & alertTime &"  - Maintenant : " & Now , 5, "Fab_Alerte",  4096 +64
			
        End If
        On Error GoTo 0
    Loop Until isTimeValid
	nowAlertTime = DateAdd("n", -1, alertTime) ' Configurer l'alerte pour 1 minutes avant l'heure entr�e
    preAlertTime = DateAdd("n", -5, alertTime) ' Configurer l'alerte pour 5 minutes avant l'heure entr�e

    Do While Now < nowAlertTime'alertTime
        If Now >= preAlertTime Then
         

            minutesRemaining = DateDiff("n", Now, alertTime)
			message.Speak "Bien, annonce FabAlert : Ceci est un rappel! Il reste maintenant " & minutesRemaining & "minutes avant l'heure de votre alerte."
            'MsgBox "Ceci est un rappel! Il reste maintenant " & minutesRemaining & " minutes avant l'heure de votre alerte.", vbOKOnly + vbInformation, " Fab_Alerte Notification"
            oShell.Popup "Ceci est un rappel! Il reste maintenant " & minutesRemaining & " minutes avant l'heure de votre alerte.", 5, "Fab_Alerte",  4096 +16
			
			CustomSleep(60) ' Attendre 60 secondes avec v�rification
        Else
            CustomSleep(30) ' V�rifier toutes les 30 secondes
        End If
    Loop
		message.Speak "Bien, annonce FabAlert : L'heure de l'alerte est pass�e. Veuillez entrer une nouvelle heure pour la prochaine alerte"
        'MsgBox "L'heure de l'alerte est pass�e. Veuillez entrer une nouvelle heure pour la prochaine alerte.", vbOKOnly + vbInformation, "Fab_Alerte Alerte Termin�e"
		oShell.Popup "L'heure de l'alerte est pass�e. Veuillez entrer une nouvelle heure pour la prochaine alerte.", 5, "Fab_Alerte",  4096 +16
		
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
				message.Speak "Bien, annonce FabAlert : Une autre instance du programme est d�tect�e. Ce programme va s'arr�ter apr�s 10 secondes."
              ' WScript.Echo "Une autre instance du programme est d�tect�e. Ce programme va s'arr�ter apr�s 10 secondes."
			    oShell.Popup "Une autre instance du programme est d�tect�e. Ce programme va s'arr�ter apr�s 10 secondes.", 5, "Fab_Alerte Programme Termin�",  4096 +16
                
				WScript.Sleep(10000)  ' Attendre 10 secondes avec v�rification
                WScript.Quit
 
End Sub
