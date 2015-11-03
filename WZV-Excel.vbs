Version = "1.05"
On Error Resume Next
url = "https://raw.githubusercontent.com/michiil/vbs_scrips/master/WZV-Excel.vbs"
Set objReq = CreateObject("Msxml2.XMLHttp.6.0")
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
    MsgBox "Update durchgefuehrt! Bitte neu starten."
    WScript.Quit
  End If
End If

If WScript.Arguments.Count = 0 Then
  MsgBox "Zum starten bitte eine Datei auf das Script ziehen."
  WScript.Quit
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile (WScript.Arguments.Item(0) , 1)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ScreenUpdating = False
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
objExcel.ScreenUpdating = True

'Konstanten: http://woonjas.linuxnerd.org/web/download.nsf/files/8C6C00FB633BCDC0C1256F39001D899E/$file/msoffice_constants.txt
