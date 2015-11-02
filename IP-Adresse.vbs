Version = "1.00"

url = "https://raw.githubusercontent.com/michiil/vbs_scrips/master/IP-Adresse.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", url, False
req.send
If req.Status = 200 Then
  ArrGit = Split(req.responseText, vbLf)
  MyOwn = Wscript.ScriptFullName
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objTextFile = objFSO.OpenTextFile(MyOwn, 1) '1 = For Reading
  ArrLocal = Split(objTextFile.ReadAll, vbCrLf)
  objTextFile.Close
  If ArrGit(0) <> ArrLocal(0) Then
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 2) '2 = For Writing
    objTextFile.Write (Join(ArrGit, vbCrLf))
    objTextFile.Close
    MsgBox "Update durchgefuehrt! Bitte neu starten."
    WScript.Quit
  End If
End If

'Variablen definieren
Dim Adapter, text, Adapternr, n, found, aproxy, regArray, switch, IP, SubNM
strComputer = "." ' . = lokaler pc
Adapter = "LAN-Verbindung"
const HKEY_CURRENT_USER = &H80000001
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
'Array definieren und Laenge setzen (kein Dim noetig)
ReDim AdapterArray(-1)
'Funktionen setzen
Set re = New RegExp
Set objIE = CreateObject("InternetExplorer.Application") 'InternetExplorer um da Langer&Laumann Webinterface zu starten.
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv") 'Fuer regedit
Set objShell = WScript.CreateObject("WScript.Shell") 'CMD Shell
Set objWMIService = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
'Rexex Funktion definieren.
With re
  .Pattern    = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$"
  .IgnoreCase = False
  .Global     = False
End With
'Regedit Funktion (Ein- und ausschalten der automatischen Proxykonfiguration)
Function autoproxy(switch)
  'Key auslesen und in Array schreiben
  objReg.GetBinaryValue HKEY_CURRENT_USER, strKeyPath, "DefaultConnectionSettings", regArray
  'Bit je nach Option beschreiben (9 = an; 1 = aus)
  If switch = "on" Then
    regArray(8) = 9
  ElseIf switch = "off" Then
    regArray(8) = 1
  Else
    MsgBox "Funktion falsch aufgerufen. (Wert " & switch & ")",0,"IP-Adresse"
  End If
  'Key zurueck in die Reg schreiben
  objReg.SetBinaryValue HKEY_CURRENT_USER, strKeyPath, "DefaultConnectionSettings", regArray
End Function
'Netzwerkadapter auslesen
For Each objItem in colItems
	If Len(objItem.NetConnectionID) Then
		ReDim Preserve AdapterArray (UBound(AdapterArray) + 1)
		AdapterArray(UBound(AdapterArray)) = objItem.NetConnectionID
	End If
Next
'Funktion fuer Netzwerkadapter auswahl.
Function netzadapt()
  'Netzwerkadapter in Array schreiben
  for n = 0 to ubound(AdapterArray)
    text = text & (n + 1) & " = " & AdapterArray(n) & VbCrLf
  Next
  'Eingabe Box
  Adapternr=InputBox("Adapter waehlen:" & VbCrLf & VbCrLf & text,"IP-Adresse")
  'Eingabe pruefen
  If (CInt(Adapternr) > CInt(UBound(AdapterArray) + 1)) OR (CInt(Adapternr) < 1) Then
    MsgBox "Ungueltige Eingabe!",0,"IP-Adresse"
  Else
    'Neuen Adapter in das Script schreiben
    MyOwn = Wscript.ScriptFullName
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 1)
    ArrAllText = Split(objTextFile.ReadAll, vbCrLf)
    objTextFile.Close
    ArrAllText(25) = "Adapter = """ & AdapterArray((Adapternr-1)) & """"
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 2)
    objTextFile.Write (Join(ArrAllText, vbCrLf))
    objTextFile.Close
  End if
  MsgBox "Adapter wurde geaendert",0,"IP-Adresse"
End Function
'Pruefen ob momentan gewaelter Adapter existiert
for n = 0 to ubound(AdapterArray)
    if AdapterArray(n) = Adapter then
        found = true
    end if
next
if found = true then
	'Eingabe Box
	Input=InputBox("Was soll gemacht werden?"&VbCRLF&VbCRLF&_
	"1 = DHCP (Firmennetz, Siemens -X127)"&VbCRLF&_
	"2 = Div feste IP's (Fanuc, MCU)"&VbCRLF&_
	"3 = Langer & Laumann Tuerautomatik"&VbCRLF&_
	"      (automatische Proykonfiguration deaktiviert)"&VbCRLF&_
	"4 = Manuell (feste IP)"&VbCRLF&_
	"5 = Netzwerkadapter aendern"&VbCRLF&_
	"      (aktuell = " & Adapter & ")"&VbCRLF&_
  "9 = Info","IP-Adresse")
	Select Case Input
	Case "1"
    'Automatische Proxy konfiguration aktivieren
    call autoproxy("on")
		'DHCP aktivieren
		objShell.Run "netsh interface ipv4 set address " & Adapter & " dhcp", 0, True
		MsgBox "DHCP Eingestellt und automatische Proxykonfiguration aktiviert.",0,"IP-Adresse"
	Case "2"
		'Diverse Feste IP's setzen
		objShell.Run "netsh interface ipv4 set address " & Adapter & " static 192.168.100.20 255.255.255.0", 0, True
		objShell.Run "netsh interface ipv4 add address " & Adapter & " 193.46.5.183 255.255.255.0", 0, True
		objShell.Run "netsh interface ipv4 add address " & Adapter & " 193.46.6.183 255.255.255.0", 0, True
		objShell.Run "netsh interface ipv4 add address " & Adapter & " 192.168.0.2 255.255.255.0", 0, True
		MsgBox "Diverse fixe IP Adressen wurden festgelegt.",0,"IP-Adresse"
	Case "3"
		'Automatische Proxy konfiguration deaktivieren
    call autoproxy("off")
		'Feste IP's setzen
		objShell.Run "netsh interface ipv4 set address " & Adapter & " static 172.16.1.151 255.255.255.0", 0, True
		MsgBox "Die IP fuer die Tuerautomaktik wurde festgelegt und die automatische Proxykonfiguration wurde deaktiviert."&VbCRLF&_
    "Das Webinterface wird jetzt gestartet.",0,"IP-Adresse"
    'InternetExplorer starten und zum Webinterface navigieren.
    objIE.Visible = 1
    objIE.Navigate "http://172.16.1.150/"
	Case "4"
    'Manuelle IP
		'Eingabe Boxen
		IP=InputBox("IP Eingeben:"&VbCRLF&VbCRLF&_
		"z.B. 193.46.8.53","IP-Adresse")
		If re.Test( IP ) Then
			SubNM=InputBox("Subnetzmaske Eingeben:"&VbCRLF&VbCRLF&_
			"z.B. 255.255.255.0","IP-Adresse","255.255.255.0")
			If re.Test( SubNM ) Then
				aproxy=MsgBox("Soll die automatische Proxykonfiguration deaktiviert werden?",4,"IP-Adresse")
				If aproxy = "6" Then
				  call autoproxy("off")
				End If
				'Manuelle IP setzen
				objShell.Run "netsh interface ipv4 set address " & Adapter & " static " & IP & " " & SUBMN, 0, True
				If aproxy = "6" Then
				  MsgBox "Die IP " & IP & " und die Subnetzmaske " & SubNM & " wurden festgelegt und die automatische Proxykonfiguration wurde deaktiviert.",0,"IP-Adresse"
				Else
				  MsgBox "Die IP " & IP & " und die Subnetzmaske " & SubNM & " wurden festgelegt.",0,"IP-Adresse"
				End If
			Else
				MsgBox "Ungueltige Subnetzmaske!",0,"IP-Adresse"
			End If
		Else
			MsgBox "Ungueltige IP!",0,"IP-Adresse"
		End If
	Case "5"
		'Netzwerkadapter aendern.
    call netzadapt()
  Case "9"
  	'Info
    MsgBox "IP-Adressen Script by Michi Lehenauer" & vbCrLf & "Version " & Version
	Case ""
		MsgBox "Abgebrochen!",0,"IP-Adresse"
	Case else
    MsgBox "Ungueltige Eingabe!",0,"IP-Adresse"
	End Select
else
  MsgBox "Der gewaelte Adapter """ & Adapter & """ existiert nicht! Bitte neuen waehlen.",0,"IP-Adresse"
	call netzadapt()
End If

WScript.Quit
