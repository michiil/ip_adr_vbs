Version = "1.00"

url = "https://raw.githubusercontent.com/michiil/vbs_scrips/master/WZV-Excel.vbs"
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

If WScript.Arguments.Count = 0 Then
  MsgBox "Zum starten bitte eine Datei auf das Scipt ziehen."
  WScript.Quit
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile (WScript.Arguments.Item(0) , 1)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
n = 0
Do Until objTextFile.AtEndOfStream
    strLine = objTextFile.Readline
    arrtext = Split(strLine , ";")
    For i = 0 to Ubound(arrtext)
		objExcel.Cells((n + 1), (i + 1)).Value = arrtext(i)
    Next
	n = n + 1
Loop

objExcel.Columns("A:D").HorizontalAlignment = -4131 'links

Function minmax(range)
	objExcel.Range(range).Select
	Set objSelection = objExcel.Selection
	objSelection.FormatConditions.Delete
	objSelection.FormatConditions.Add 1, 3, "=MAX(" & range & ")" '1=xlcellvalue; 3=xlequal
	objSelection.FormatConditions(1).Interior.ColorIndex = 3
	objSelection.FormatConditions.Add 1, 3, "=MIN(" & range & ")" '1=xlcellvalue; 3=xlequal
	objSelection.FormatConditions(2).Interior.ColorIndex = 4
End Function

'call minmax("$A$19:$A$24")
'call minmax("$B$19:$B$24")
'call minmax("$C$19:$C$24")
'call minmax("$D$19:$D$24")

'Konstanten: http://woonjas.linuxnerd.org/web/download.nsf/files/8C6C00FB633BCDC0C1256F39001D899E/$file/msoffice_constants.txt
