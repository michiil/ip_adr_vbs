Version = "3.00"
'V3.00 Neue Programm strukturierung
On Error Resume Next
SetLocale(1033)
url = "https://raw.githubusercontent.com/michiil/vbs_scrips/master/IP-Adresse.vbs"
Set objReq = CreateObject("Msxml2.ServerXMLHttp.6.0")
objReq.setTimeouts 500,500,500,500
objReq.open "GET", url, False
objReq.send
If Err.Number = 0 Then
  If objReq.Status = 200 Then
    ArrGit = Split(objReq.responseText, vbLf)
    MyOwn = Wscript.ScriptFullName
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 1) '1 = For Reading
    ArrLocal = Split(objTextFile.ReadAll, vbCrLf)
    objTextFile.Close
    VerLocal = Split(ArrLocal(0),"""")
    VerGit = Split(ArrGit(0),"""")
    If CSng(VerGit(1)) > CSng(VerLocal(1)) Then
      Set objTextFile = objFSO.OpenTextFile(MyOwn, 2) '2 = For Writing
      objTextFile.Write (Join(ArrGit, vbCrLf))
      objTextFile.Close
      MsgBox "Update durchgefuehrt! Bitte neu starten." & VbCRLF & ArrGit(1)
      WScript.Quit
    End If
  End If
Else
  Err.Clear
End If

Nic = "LAN-Verbindung"

Set objIE = CreateObject("InternetExplorer.Application")
Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")

Set IpRegex = New RegExp
With IpRegex
  .Pattern    = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$"
  .IgnoreCase = False
  .Global     = False
End With

Set IpPortRegex = New RegExp
With IpPortRegex
  .Pattern    = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}\:[0-9]{2,4}$"
  .IgnoreCase = False
  .Global     = False
End With

' ███████ ███████ ████████ ███    ██ ██  ██████
' ██      ██         ██    ████   ██ ██ ██
' ███████ █████      ██    ██ ██  ██ ██ ██
'      ██ ██         ██    ██  ██ ██ ██ ██
' ███████ ███████    ██    ██   ████ ██  ██████

Function SetNic()
  for n = 1 to ubound(NicArray)
    NicList = NicList & n & " = " & NicArray(n) & VbCrLf
  Next
  NicNr=InputBox("Adapter waehlen:" & VbCrLf & VbCrLf & NicList,"IP-Adresse")
  If (CInt(NicNr) > CInt(UBound(NicArray))) OR (CInt(NicNr) < 1) Then
    MsgBox "Ungueltige Eingabe!",0,"IP-Adresse"
  Else
    MyOwn = Wscript.ScriptFullName
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 1)
    ArrAllText = Split(objTextFile.ReadAll, vbCrLf)
    objTextFile.Close
    ArrAllText(31) = "Nic = """ & NicArray(NicNr) & """"
    Set objTextFile = objFSO.OpenTextFile(MyOwn, 2)
    objTextFile.Write (Join(ArrAllText, vbCrLf))
    objTextFile.Close
    MsgBox "Adapter wurde auf " & NicArray(NicNr) & " geaendert",0,"IP-Adresse"
    Nic = NicArray(NicNr)
  End if
End Function

'  ██████ ██   ██ ███████  ██████ ██   ██ ███    ██ ██  ██████
' ██      ██   ██ ██      ██      ██  ██  ████   ██ ██ ██
' ██      ███████ █████   ██      █████   ██ ██  ██ ██ ██
' ██      ██   ██ ██      ██      ██  ██  ██  ██ ██ ██ ██
'  ██████ ██   ██ ███████  ██████ ██   ██ ██   ████ ██  ██████

ReDim NicArray(0)
For Each objItem in colItems
  If Len(objItem.NetConnectionID) Then
    ReDim Preserve NicArray (UBound(NicArray) + 1)
    NicArray(UBound(NicArray)) = objItem.NetConnectionID
  End If
Next
For n = 0 to ubound(NicArray)
  If NicArray(n) = Nic then
    NicFound = true
  End If
Next
If Not NicFound = true then
  MsgBox "Der gewaelte Adapter """ & Nic & """ existiert nicht! Bitte neuen waehlen.",0,"IP-Adresse"
  call SetNic()
End If

' ██████  ██████   ██████  ██   ██ ██    ██
' ██   ██ ██   ██ ██    ██  ██ ██   ██  ██
' ██████  ██████  ██    ██   ███     ████
' ██      ██   ██ ██    ██  ██ ██     ██
' ██      ██   ██  ██████  ██   ██    ██

Function proxy(task, switch, quiet)
  proxysvr=objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
  proxyenable=objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
  proxyreg=objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections\DefaultConnectionSettings")
  If proxyreg(8) AND 2^3 Then
    autoproxy = 1
  Else
    autoproxy = 0
  End If
  Select Case task
  Case "auto"
    If switch = autoproxy Then
      If Not quiet = 1 Then
        MsgBox "Autoproxy ist schon auf " & autoproxy,0,"IP-Adresse"
      End If
      Exit Function
    End If
    If ((autoproxy = 0) And ((switch = 1) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoDetect", 1, "REG_DWORD"
      If Not quiet = 1 Then
        GoOn = MsgBox("Automatische Proxykonfiguration Eingeschaltet!" & VbCRLF & _
        "Aenderung wird beim naechsten Start von IE wirksam." & VbCRLF & _
        "Noch was aendern?",260,"IP-Adresse")
        If GoOn = 6 Then
          call proxy("ask",0,0)
        End If
      End If
    ElseIf ((autoproxy = 1) And ((switch = 0) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoDetect", 0, "REG_DWORD"
      If Not quiet = 1 Then
        GoOn = MsgBox("Automatische Proxykonfiguration Ausgeschaltet!" & VbCRLF & _
        "Aenderung wird beim naechsten Start von IE wirksam." & VbCRLF & _
        "Noch was aendern?",260,"IP-Adresse")
        If GoOn = 6 Then
          call proxy("ask",0,0)
        End If
      End If
    Else
      MsgBox "Error!",0,"IP-Adresse"
    End If
  Case "proxy"
    If switch = proxyenable Then
      If Not quiet = 1 Then
        MsgBox "Proxy ist schon auf " & proxyenable,0,"IP-Adresse"
      End If
      Exit Function
    End If
    If ((proxyenable = 0) And ((switch = 1) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
      If Not quiet = 1 Then
        GoOn = MsgBox("Proxy Eingeschaltet!" & VbCRLF & _
        "Noch was aendern?",260,"IP-Adresse")
        If GoOn = 6 Then
          call proxy("ask",0,0)
        End If
      End If
    ElseIf ((proxyenable = 1) And ((switch = 0) Or (switch = 2))) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
      If Not quiet = 1 Then
        GoOn = MsgBox("Proxy Ausgeschaltet!" & VbCRLF & _
        "Noch was aendern?",260,"IP-Adresse")
        If GoOn = 6 Then
          call proxy("ask",0,0)
        End If
      End If
    Else
      MsgBox "Error!",0,"IP-Adresse"
    End If
  Case "proxysvr"
    NewProxy=Inputbox("Neuen Proxyserver mit Port eingeben:" & VbCRLF & _
    "(z.B. 192.164.2.20:8080)","IP-Adresse")
    If IpPortRegex.Test( NewProxy ) Then
      objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", NewProxy, "REG_SZ"
      GoOn = MsgBox("Proxyserver wurde auf " & NewProxy & " geaendert!" & VbCRLF & _
      "Noch was aendern?",260,"IP-Adresse")
      If GoOn = 6 Then
        call proxy("ask",0,0)
      End If
    Else
      MsgBox "Ungueltiges Eingabeformat!",0,"IP-Adresse"
    End If
  Case "ask"
    ProxySet=InputBox("Die momentanen Einstellungen sind:" & VbCRLF & VbCRLF & _
    "Proxyserver:                " & proxysvr & VbCRLF & _
    "Proxy:                          " & proxyenable & VbCRLF & _
    "Automatische Proxy:    " & autoproxy & VbCRLF & VbCRLF & _
    "Was soll geaendert werden?" & VbCRLF & _
    "1 = Proxyserver" & VbCRLF & _
    "2 = Proxy Ein- bzw. Ausschalten" & VbCRLF & _
    "3 = Automatische Proxykonfiguration" & VbCRLF,"IP-Adresse")
    Select Case ProxySet
    Case "1"
      call proxy("proxysvr",2,0)
    Case "2"
      call proxy("proxy",2,0)
    Case "3"
      call proxy("auto",2,0)
    End Select
  Case else
    MsgBox "error proxy",0,"IP-Adresse"
  End Select
End Function

' ███████ ███████ ████████ ██ ██████
' ██      ██         ██    ██ ██   ██
' ███████ █████      ██    ██ ██████
'      ██ ██         ██    ██ ██
' ███████ ███████    ██    ██ ██

Function setIP(task,IpArray)
  Select Case task
  Case "static"
    objShell.Run "netsh interface ipv4 set address """ & Nic & """ static " & IpArray(0), 0, True
    if UBound(IpArray) > 0 then
      For n = 1 to ubound(IpArray)
        objShell.Run "netsh interface ipv4 add address """ & Nic & """ " & IpArray(n), 0, True
      Next
    End If
  Case "dhcp"
    objShell.Run "netsh interface ipv4 set address """ & Nic & """ dhcp", 0, True
  Case "reset"
    objShell.Run "netsh interface set interface """ & Nic & """ disabled", 0, True
    objShell.Run "netsh interface set interface """ & Nic & """ enabled", 0, True
  Case else
    MsgBox "error setIP",0,"IP-Adresse"
  End Select
End Function

' ███    ███  █████  ██ ███    ██     ███    ███ ███████ ███    ██ ██    ██
' ████  ████ ██   ██ ██ ████   ██     ████  ████ ██      ████   ██ ██    ██
' ██ ████ ██ ███████ ██ ██ ██  ██     ██ ████ ██ █████   ██ ██  ██ ██    ██
' ██  ██  ██ ██   ██ ██ ██  ██ ██     ██  ██  ██ ██      ██  ██ ██ ██    ██
' ██      ██ ██   ██ ██ ██   ████     ██      ██ ███████ ██   ████  ██████

Input=InputBox("Was soll gemacht werden?" & VbCRLF & VbCRLF & _
"1 = DHCP (Firmennetz, Siemens -X127)" & VbCRLF & _
"2 = Div feste IP's (Fanuc, MCU)" & VbCRLF & _
"3 = Langer & Laumann Tuerautomatik" & VbCRLF & _
"4 = Manuell (feste IP)" & VbCRLF & _
"5 = Netzwerkadapter aendern" & VbCRLF & _
"      (aktuell = " & Nic & ")" & VbCRLF & _
"6 = Proxyeinstellungen" & VbCRLF & _
"7 = Ethernet Neustart (fuer 828D)" & VbCRLF & _
"9 = Info" & VbCRLF,"IP-Adresse")
Select Case Input
Case "1"
  call setIP("dhcp",0)
  DHCPProxy=MsgBox("DHCP Eingestellt." & VbCRLF & VbCRLF & _
  "Sollen die Proxyeinstellungen geaendert werden?",4,"IP-Adresse")
  If DHCPProxy = 6 Then
    call proxy("ask",0,0)
  End If
Case "2"
  call setIP("static",Array("192.168.100.20 255.255.255.0","193.46.5.183 255.255.255.0","193.46.6.183 255.255.255.0","192.168.0.2 255.255.255.0","192.168.214.30 255.255.255.0"))
  MsgBox "Folgende IP Adressen wurden festgelegt:" & VbCRLF & VbCRLF & _
  "192.168.100.20 255.255.255.0 (Fanuc Ethernet)" & VbCRLF & _
  "193.46.5.183 255.255.255.0 (Fanuc Ethernet)" & VbCRLF & _
  "193.46.6.183 255.255.255.0 (Fanuc Ethernet)" & VbCRLF & _
  "192.168.214.30 255.255.255.0 (Siemens PCU50)" & VbCRLF & _
  "192.168.0.2 255.255.255.0 (Visualisierung MCU)",0,"IP-Adresse"
Case "3"
  call setIP("static",Array("172.16.1.151 255.255.255.0"))
  LLProxy=MsgBox("Die IP fuer die Tuerautomaktik wurde festgelegt." & VbCRLF & _
  "Das Webinterface wird jetzt gestartet. Soll noch die Proxy deaktiviert werden?",4,"IP-Adresse")
  If LLProxy = 6 Then
    call proxy("proxy",0,1)
    call proxy("auto",0,1)
    WScript.Sleep 1500
  End If
  objIE.Visible = 1
  objIE.Navigate "http://172.16.1.150/"
Case "4"
  IP=InputBox("IP Eingeben:" &  VbCRLF & VbCRLF & _
  "z.B. 193.46.8.53","IP-Adresse")
  If IpRegex.Test( IP ) Then
    SubNM=InputBox("Subnetzmaske Eingeben:" & VbCRLF & VbCRLF & _
    "z.B. 255.255.255.0","IP-Adresse","255.255.255.0")
    If IpRegex.Test( SubNM ) Then
      call setIP("static",Array(IP & " " & SUBMN))
      ManProxy=MsgBox("Die IP " & IP & " und die Subnetzmaske " & SubNM & " wurden festgelegt." & VbCRLF & VbCRLF & _
      "Sollen die Proxyeinstellungen geaendert werden?",4,"IP-Adresse")
      If ManProxy = 6 Then
        call proxy("ask",0,0)
      End If
    Else
      MsgBox "Ungueltige Subnetzmaske!",0,"IP-Adresse"
    End If
  Else
    MsgBox "Ungueltige IP!",0,"IP-Adresse"
  End If
Case "5"
  call SetNic()
Case "6"
  call proxy("ask",0,0)
Case "7"
  call setIP("reset",0)
Case "9"
  MsgBox "IP-Adressen Script by Michi Lehenauer" & VbCRLF & _
  "Version " & Version,0,""
Case ""
Case else
  MsgBox "Ungueltige Eingabe!",0,"IP-Adresse"
End Select

WScript.Quit
