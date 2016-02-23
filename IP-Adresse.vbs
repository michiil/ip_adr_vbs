Version = "2.00"
'V2.00 Proxy Einstellungen.
On Error Resume Next

url = "https://raw.githubusercontent.com/michiil/vbs_scrips/master/IP-Adresse.vbs"
Set objReq = CreateObject("Msxml2.ServerXMLHttp.6.0")
objReq.setTimouts 5000,5000,5000,5000
objReq.open "GET", url, False
objReq.send
If objReq.Status = 200 Then
  ArrGit = Split(objReq.responseText, vbLf)
  MyOwn = Wscript.ScriptFullName
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objTextFile = objFSO.OpenTextFile(MyOwn, 1) '1 = For Reading
  ArrLocal = Split(objTextFile.ReadAll, vbCrLf)
  objTextFile.Close
  If ArrGit(0) <> ArrLocal(0) Then
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 2) '2 = For Writing
    objTextFile.Write (Join(ArrGit, vbCrLf))
    objTextFile.Close
    MsgBox "Update durchgefuehrt! Bitte neu starten." & VbCRLF & ArrGit(1)
    WScript.Quit
  End If
End If

Adapter = "LAN-Verbindung"
'Dynamischen Array mit Laenge 0 definieren
ReDim AdapterArray(0)
'Funktionen setzen
Set IpRegex = New RegExp
Set IpPortRegex = New RegExp
Set objIE = CreateObject("InternetExplorer.Application") 'InternetExplorer um das Langer&Laumann Webinterface zu starten.
Set objShell = WScript.CreateObject("WScript.Shell") 'CMD Shell
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2") 'Netzwerkadapter auslesen
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter") 'Netzwerkadapter auslesen
'Rexex fuer IP ueberpruefung definieren.
With IpRegex
  .Pattern    = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$"
  .IgnoreCase = False
  .Global     = False
End With
'Rexex fuer IP:Port ueberpruefung definieren.
With IpPortRegex
  .Pattern    = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}\:[0-9]{2,4}$"
  .IgnoreCase = False
  .Global     = False
End With
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
  for n = 1 to ubound(AdapterArray)
    text = text & n & " = " & AdapterArray(n) & VbCrLf
  Next
  'Eingabe Box
  Adapternr=InputBox("Adapter waehlen:" & VbCrLf & VbCrLf & text,"IP-Adresse")
  'Eingabe pruefen
  If (CInt(Adapternr) > CInt(UBound(AdapterArray))) OR (CInt(Adapternr) < 1) Then
    MsgBox "Ungueltige Eingabe!",0,"IP-Adresse"
  Else
    'Neuen Adapter in das Script schreiben
    MyOwn = Wscript.ScriptFullName
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 1)
    ArrAllText = Split(objTextFile.ReadAll, vbCrLf)
    objTextFile.Close
    ArrAllText(25) = "Adapter = """ & AdapterArray(Adapternr) & """"
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 2)
    objTextFile.Write (Join(ArrAllText, vbCrLf))
    objTextFile.Close
    MsgBox "Adapter wurde auf " & AdapterArray(Adapternr) & " geaendert",0,"IP-Adresse"
  End if
End Function

'Proxyfunktion
Function proxy(task, switch, quiet)
  proxysvr=objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
  proxyenable=objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
  proxyreg=objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections\DefaultConnectionSettings")
  If proxyreg(8) >= 9 Then
    autoproxy = 1
  Else
    autoproxy = 0
  End If
  Select Case task
  Case "auto"
    If switch = autoproxy Then
      If Not quiet = 1 Then
        MsgBox "Autoproxy ist schon auf " & autoproxy,0,"Proxyeinstellungen"
      End If
      Exit Function
    End If
    If ((autoproxy = 0) And ((switch = 1) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoDetect", 1, "REG_DWORD"
      If Not quiet = 1 Then
        MsgBox "Automatische Proxykonfiguration Eingeschaltet!" & VbCRLF & _
        "Aenderung wird beim naechsten Start von IE wirksam.",0,"Proxyeinstellungen"
      End If
    ElseIf ((autoproxy = 1) And ((switch = 0) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoDetect", 0, "REG_DWORD"
      If Not quiet = 1 Then
        MsgBox "Automatische Proxykonfiguration Ausgeschaltet!" & VbCRLF & _
        "Aenderung wird beim naechsten Start von IE wirksam.",0,"Proxyeinstellungen"
      End If
    Else
      MsgBox "Error!",0,"Proxyeinstellungen"
    End If
  Case "proxy"
    If switch = proxyenable Then
      If Not quiet = 1 Then
        MsgBox "Proxy ist schon auf " & proxyenable,0,"Proxyeinstellungen"
      End If
      Exit Function
    End If
    If ((proxyenable = 0) And ((switch = 1) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
      If Not quiet = 1 Then
        MsgBox "Proxy Eingeschaltet!",0,"Proxyeinstellungen"
      End If
    ElseIf ((proxyenable = 1) And ((switch = 0) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
      If Not quiet = 1 Then
        MsgBox "Proxy Ausgeschaltet!",0,"Proxyeinstellungen"
      End If
    Else
      MsgBox "Error!",0,"Proxyeinstellungen"
    End If
  Case "proxysvr"
    NewProxy=Inputbox("Neuen Proxyserver mit Port eingeben:" & VbCRLF & _
    "(z.B. 192.164.2.20:8080)","Proxyeinstellungen")
    If IpPortRegex.Test( NewProxy ) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", NewProxy, "REG_SZ"
      MsgBox "Proxyserver wurde auf " & NewProxy & " geaendert!",0,"Proxyeinstellungen"
    Else
      MsgBox "Ungueltiges Eingabeformat!",0,"IP-Adresse"
    End If
  Case "return"
    proxy="Proxyserver:                " & proxysvr & VbCRLF & _
    "Proxy:                          " & proxyenable & VbCRLF & _
    "Automatische Proxy:    " & autoproxy
  End Select
End Function

'Pruefen ob momentan gewaelter Adapter existiert
For n = 0 to ubound(AdapterArray)
  If AdapterArray(n) = Adapter then
    found = true
  End If
next
If Not found = true then
  MsgBox "Der gewaelte Adapter """ & Adapter & """ existiert nicht! Bitte neuen waehlen.",0,"IP-Adresse"
  call netzadapt()
End If

'Eingabe Box
Input=InputBox("Was soll gemacht werden?" & VbCRLF & VbCRLF & _
"1 = DHCP (Firmennetz, Siemens -X127)" & VbCRLF & _
"2 = Div feste IP's (Fanuc, MCU)" & VbCRLF & _
"3 = Langer & Laumann Tuerautomatik" & VbCRLF & _
"4 = Manuell (feste IP)" & VbCRLF & _
"5 = Netzwerkadapter aendern" & VbCRLF & _
"      (aktuell = " & Adapter & ")" & VbCRLF & _
"6 = Proxyeinstellungen" & VbCRLF & _
"9 = Info" & VbCRLF,"IP-Adresse")
Select Case Input
Case "1" 'DHCP
  objShell.Run "netsh interface ipv4 set address """ & Adapter & """ dhcp", 0, True
  proxyreturn = proxy("return",0,0)
  DHCPProxy=InputBox("DHCP Eingestellt." & VbCRLF & VbCRLF & _
  "Momentane Proxyeinstellungen:" & VbCRLF & VbCRLF & _
  proxyreturn & VbCRLF & VbCRLF & _
  "1 = Proxy aktivieren" & VbCRLF & _
  "2 = Automatische Proxykonfiguration aktivieren" & VbCRLF & _
  "3 = Proxy aktivieren und Proxyserver bearbeiten" & VbCRLF & _
  "Alles Andere = Nichts tun" & VbCRLF,"IP-Adresse")
  Select Case DHCPProxy
  Case "1"
    call proxy("proxy",1,0)
  Case "2"
    call proxy("auto",1,0)
  Case "3"
    call proxy("proxysvr",0,0)
    call proxy("proxy",1,0)
  End Select
Case "2" 'Diverse Feste IP's setzen
  objShell.Run "netsh interface ipv4 set address """ & Adapter & """ static 192.168.100.20 255.255.255.0", 0, True
  objShell.Run "netsh interface ipv4 add address """ & Adapter & """ 193.46.5.183 255.255.255.0", 0, True
  objShell.Run "netsh interface ipv4 add address """ & Adapter & """ 193.46.6.183 255.255.255.0", 0, True
  objShell.Run "netsh interface ipv4 add address """ & Adapter & """ 192.168.0.2 255.255.255.0", 0, True
  MsgBox "Folgende IP Adressen wurden festgelegt:" & VbCRLF & VbCRLF & _
  "192.168.100.20 255.255.255.0 (Fanuc Ethernet)" & VbCRLF & _
  "193.46.5.183 255.255.255.0 (Fanuc Ethernet)" & VbCRLF & _
  "193.46.6.183 255.255.255.0 (Fanuc Ethernet)" & VbCRLF & _
  "192.168.0.2 255.255.255.0 (Visualisierung MCU)",0,"IP-Adresse"
Case "3" 'Langer & Laumann Tuerautomatik
  'Feste IP's setzen
  objShell.Run "netsh interface ipv4 set address """ & Adapter & """ static 172.16.1.151 255.255.255.0", 0, True
  LLProxy=MsgBox("Die IP fuer die Tuerautomaktik wurde festgelegt." & VbCRLF & _
  "Das Webinterface wird jetzt gestartet. Soll noch die Proxy deaktiviert werden?",4,"IP-Adresse")
  If LLProxy = 6 Then
    call proxy("proxy",0,1)
    call proxy("auto",0,1)
    WScript.Sleep 500
  End If
  'InternetExplorer starten und zum Webinterface navigieren.
  objIE.Visible = 1
  objIE.Navigate "http://172.16.1.150/"
Case "4" 'Manuelle IP
  'Eingabe Boxen
  IP=InputBox("IP Eingeben:" &  VbCRLF & VbCRLF & _
  "z.B. 193.46.8.53","IP-Adresse")
  If IpRegex.Test( IP ) Then
    SubNM=InputBox("Subnetzmaske Eingeben:" & VbCRLF & VbCRLF & _
    "z.B. 255.255.255.0","IP-Adresse","255.255.255.0")
    If IpRegex.Test( SubNM ) Then
      'Manuelle IP setzen
      objShell.Run "netsh interface ipv4 set address """ & Adapter & """ static " & IP & " " & SUBMN, 0, True
      MsgBox "Die IP " & IP & " und die Subnetzmaske " & SubNM & " wurden festgelegt.",0,"IP-Adresse"
    Else
      MsgBox "Ungueltige Subnetzmaske!",0,"IP-Adresse"
    End If
  Else
    MsgBox "Ungueltige IP!",0,"IP-Adresse"
  End If
Case "5" 'Netzwerkadapter aendern.
  call netzadapt()
Case "6" 'Proyeinstellungen.
  proxyreturn = proxy("return",0,0)
  ProxySet=InputBox("Die momentanen Einstellungen sind:" & VbCRLF & VbCRLF & _
  proxyreturn & VbCRLF & VbCRLF & _
  "Was soll geaendert werden?" & VbCRLF & _
  "1 = Proxyserver" & VbCRLF & _
  "2 = Proxy Ein- bzw. Ausschalten" & VbCRLF & _
  "3 = Automatische Proxykonfiguration" & VbCRLF,"Proxyeinstellungen")
  Select Case ProxySet
  Case "1"
    call proxy("proxysvr",2,0)
  Case "2"
    call proxy("proxy",2,0)
  Case "3"
    call proxy("auto",2,0)
  End Select
Case "9" 'Info
  MsgBox "IP-Adressen Script by Michi Lehenauer" & VbCRLF & _
  "Version " & Version,0,""
Case ""
  MsgBox "Abgebrochen!",0,"IP-Adresse"
Case else
  MsgBox "Ungueltige Eingabe!",0,"IP-Adresse"
End Select

WScript.Quit
