'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
'
'	Codes retour :
'		0 - Succes
'		3 - Erreur lors de la copie du répertoire Nagios
'		4 - Erreur lors de l'installation du client Nagios
'		5 - Erreur lors du démarrage du service
'		6 - Erreur lors de la création du fichier de log
'		7 - Erreur lors de l'accès au fichier de log
'
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

'	Déclaration des constantes
Const ForReading = 1, ForWriting = 2, ForAppending = 8

'	Déclaration des objets
Dim oShell, Fso, objWMIService, colItems, colListOfServices, objNewJob

'	Définition des objets
Set oShell = CreateObject("Wscript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

'	Déclaration des variables
Dim ScriptVer, strComputer, appPath, logPath, strFile, strIP, strVar
Dim progPath, NomScript, FichierLog, ActivEnteteLog, ActivEndLog

'	Définition des varaibles
ScriptVer = "1.0"
strComputer = oShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
windir = oShell.ExpandEnvironmentStrings("%windir%")
appPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")-1)
logPath = appPath & "\log"
strFile = appPath & "\Fichier\infos"
strIP = "10.150.20"
strService = "NSClientpp"
NomScript = Left(WScript.ScriptName,InStr(WScript.ScriptName,".") - 1)
FichierLog = logPath & "\" & strComputer & ".log"
ActivEnteteLog = True
ActivEndLog = False

'
'	Debut du programme
'

' Recuperation de l'adresse IP
Call TraceLog(1,"Verification adresse IP")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objIP = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

For Each IP in objIP
	j=j+1
	myCount = objIP.count
	MyArray = Split(IP.IPAddress(i),".")
	MyValue = MyArray(0) & "." & MyArray(1) & "." & MyArray(2)
	If MyValue = strIP Then
		Exit For
	Else
		If myCount = j Then
			WScript.Quit
		End If
	End if
Next

Set objIP = Nothing
Set objWMIService = Nothing


' Detection architecture
If FOsArchitecture = "64" Then
	progPath = "C:\Program Files (x86)"
	Call TraceLog(1,"Architecture de l'OS : 64 bits")
Else
	progPath = "C:\Program Files"
	Call TraceLog(1,"Architecture de l'OS : 32 bits")
End If

' Verification de presence de l'exe
If Not Fso.FileExists(progPath & "\NSClient\NSClient++.exe") Then
	Call TraceLog(1,"Le client n'existe pas sur le poste, copie des fichiers")
	' Copie de l'agent NSClient
	On Error Resume Next
	Fso.CopyFolder appPath & "\NSClient", progPath & "\NSClient"
	If Err <> 0 Then
		ActivEnteteLog = False
		ActivEndLog = True
		Call TraceLog(3,"Erreur pendant la copie des fichiers")
		WScript.Quit (3)
	End If
	shortProgPath = Fso.GetFolder(progPath & "\NSClient").ShortPath
	
	' Installation du client Nagios
	oShell.Run "cmd /c " & shortProgPath & "\NSClient++.exe /install",0,True
	If err <> 0 Then
		ActivEnteteLog = False
		ActivEndLog = True
		Call TraceLog(3,"Erreur pendant l'installation du client")
		Call TraceLog(3,Err.Description & " : " & Err.Number)
		WScript.Quit (4)
	Else
		Call TraceLog(1,"Installation du service reussie")
	End If
	
	'	Installation du systray NSClient
	oShell.Run "cmd /c " & shortProgPath & "\NSClient++.exe SysTray /install",0,True
	If err <> 0 Then
		Call TraceLog(3,"Erreur pendant l'installation du systray")
		Call TraceLog(3,Err.Description & " : " & Err.Number)
	Else
		Call TraceLog(1,"Installation du systray reussie")
	End If
	
	'	Vérification du service
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service where Name = '" & strService & "'")
	
	For Each objService in colListOfServices
		If objService.DesktopInteract = False Then
			oShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\services\NSClientpp\Type",272,"REG_DWORD"
			If err <> 0 Then
				Call TraceLog(3,"Erreur pendant la modification du registre")
				Call TraceLog(3,Err.Description & " : " & Err.Number)
			Else
				Call TraceLog(1,"Modification du registre reussie")
			End If
		End If
	Next
	
	oShell.Run "cmd /c " & shortProgPath & "\nsclient++ /start",0,True
	If err <> 0 Then
		ActivEnteteLog = False
		ActivEndLog = True		
		Call TraceLog(3,"Erreur pendant le démarrage du service")
		Call TraceLog(3,Err.Description & " : " & Err.Number)
		WScript.Quit (5)
	Else
		Call TraceLog(1,"Demarrage du service reussi")
	End If
Else
	Call TraceLog(1,"Le client est deja installe.")
	WScript.Quit
End If

' Creation de la tache planifiée pour fichier log
oShell.Run "schtasks.exe /create /SC weekly /D MON /TN LogNsClient /TR " & shortProgPath & "\checklog.vbs /ST 12:00 /RU System"
If err <> 0 Then
		ActivEnteteLog = False
		ActivEndLog = True		
		Call TraceLog(3,"Erreur pendant la creation de la tache planifiee")
		Call TraceLog(3,Err.Description & " : " & Err.Number)
		WScript.Quit (5)
	Else
		Call TraceLog(1,"Creation de la tache planifiee reussie")
	End If

Set Fso = Nothing
Set oShell = Nothing


'
'	Fin du programme
'

'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
Function FOsArchitecture
On Error Resume Next
	Dim objWMIService, colOSys, OS
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")	
	Set colOSys = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

	FOsArchitecture = Null
	For Each OS In colOSys
		FOsArchitecture = Left(OS.OSArchitecture,2)
	Next

	Set objWMIService = Nothing
	Set colOSys = Nothing
End Function

'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
Sub TraceLog(LevelAlert,strDataToLog)
	'	Gestion du fichier de log
	On Error Resume Next
	Dim LibLevelAlert
	Err.Clear
	
	Select case LevelAlert
		case 1
			LibLevelAlert = "INFORMATION"
		case 2
			LibLevelAlert = "       WARNING"
		case 3
			LibLevelAlert = "            ERROR"
	End Select   
	
	'	Ouverture en écriture de DebugFile 
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If Not oFSO.FileExists(FichierLog) Then
		'	Création du fichier de log
		set oFile = oFSO.CreateTextFile(FichierLog)
		If Err.Number <> 0 Then
			oShell.LogEvent 1, "Nagios Client : Cannot create " & FichierLog & _
				vbCrLf & "Please, contact your IT support"
			WScript.Quit (6)
		End If
	Else
		'	Ouverture en ajout du fichier de log
		Set oFile = oFSO.OpenTextFile(FichierLog,ForAppending)
		If Err.Number <> 0 Then
			oShell.LogEvent 1, "Nagios Client : Cannot modify " & FichierLog & _
				vbCrLf & "Please, contact your IT support"
			WScript.Quit (7)
		End If
	End If
	
	If ActivEnteteLog = True Then
		oFile.Writeline(now & "  " &"--------------------------------------------------------------------------------")
		oFile.Writeline(now & "  " &" START : " & WScript.ScriptFullName )
		oFile.Writeline(now & "  " &" Computer : " & strComputer )
		oFile.Writeline(now & "  " &" Version : " & ScriptVer )
		oFile.Writeline(now & "  " &"--------------------------------------------------------------------------------")
		ActivEnteteLog = false
	End If
	
	If CInt(MinLevelAlertLevel) <= CInt(LevelAlert)  Then
		'	MAJ du contenu du fichier de trace
		oFile.Writeline(Now & " - " & LibLevelAlert & " : " & strDataToLog)
	End If

	If ActivEndLog = True Then
		oFile.Writeline(now & "  " &"--------------------------------------------------------------------------------")
		oFile.Writeline(now & "  " &" END")
		oFile.Writeline(now & "  " &"--------------------------------------------------------------------------------")
	End If
	
	'	FERMETURE du Fichier  
	oFile.Close
	Set oFSO = Nothing
	Set oFile = Nothing
End Sub