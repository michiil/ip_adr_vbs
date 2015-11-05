Version = "2.00"
'On Error Resume Next
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

'Kaliebrieren K1:
objExcel.Range("E27").Value = "Max Abweichung"
objExcel.Range("A27").Formula = "=ROUND(ABS(MAX((ABS(MAX(A21:A26)-B14)),(ABS(MIN(A21:A26)-B14)))),3)"
objExcel.Range("B27").Formula = "=ROUND(ABS(MAX((ABS(MAX(B21:B26)-B15)),(ABS(MIN(B21:B26)-B15)))),3)"
objExcel.Range("C27").Formula = "=ROUND(ABS(MAX((ABS(MAX(C21:C26)-B16)),(ABS(MIN(C21:C26)-B16)))),3)"
objExcel.Range("D27").Formula = "=ROUND(ABS(MAX((ABS(MAX(D21:D26)-B17)),(ABS(MIN(D21:D26)-B17)))),3)"
objExcel.Range("A27:E27").Font.Bold = True
objExcel.Range("E28").Value = "Wiederholgenauigkeit"
objExcel.Range("A28").Formula = "=ABS(MAX(A21:A26)-MIN(A21:A26))"
objExcel.Range("B28").Formula = "=ABS(MAX(B21:B26)-MIN(B21:B26))"
objExcel.Range("C28").Formula = "=ABS(MAX(C21:C26)-MIN(C21:C26))"
objExcel.Range("D28").Formula = "=ABS(MAX(D21:D26)-MIN(D21:D26))"
objExcel.Range("A28:E28").Font.Bold = True

'Messen K1:
objExcel.Range("E38").Value = "Max Abweichung"
objExcel.Range("A38").Formula = "=MAX(ABS(MIN(A32:A37)),MAX(A32:A37))"
objExcel.Range("B38").Formula = "=MAX(ABS(MIN(B32:B37)),MAX(B32:B37))"
objExcel.Range("C38").Formula = "=MAX(ABS(MIN(C32:C37)),MAX(C32:C37))"
objExcel.Range("D38").Formula = "=MAX(ABS(MIN(D32:D37)),MAX(D32:D37))"
objExcel.Range("A38:E38").Font.Bold = True
objExcel.Range("E39").Value = "Wiederholgenauigkeit"
objExcel.Range("A39").Formula = "=ABS(MAX(A32:A37)-MIN(A32:A37))"
objExcel.Range("B39").Formula = "=ABS(MAX(B32:B37)-MIN(B32:B37))"
objExcel.Range("C39").Formula = "=ABS(MAX(C32:C37)-MIN(C32:C37))"
objExcel.Range("D39").Formula = "=ABS(MAX(D32:D37)-MIN(D32:D37))"
objExcel.Range("A39:E39").Font.Bold = True

'Aus- und Einspannen K1:
objExcel.Range("E49").Value = "Max Abweichung"
objExcel.Range("A49").Formula = "=MAX(ABS(MIN(A43:A48)),MAX(A43:A48))"
objExcel.Range("B49").Formula = "=MAX(ABS(MIN(B43:B48)),MAX(B43:B48))"
objExcel.Range("A49:E49").Font.Bold = True
objExcel.Range("E50").Value = "Wiederholgenauigkeit"
objExcel.Range("A50").Formula = "=ABS(MAX(A43:A48)-MIN(A43:A48))"
objExcel.Range("B50").Formula = "=ABS(MAX(B43:B48)-MIN(B43:B48))"
objExcel.Range("A50:E50").Font.Bold = True

If objExcel.Range("A51").Value <> "" then
  'Kaliebrieren K2:
  objExcel.Range("E73").Value = "Max Abweichung"
  objExcel.Range("A73").Formula = "=ROUND(ABS(MAX((ABS(MAX(A67:A72)-B60)),(ABS(MIN(A67:A72)-B60)))),3)"
  objExcel.Range("B73").Formula = "=ROUND(ABS(MAX((ABS(MAX(B67:B72)-B61)),(ABS(MIN(B67:B72)-B61)))),3)"
  objExcel.Range("C73").Formula = "=ROUND(ABS(MAX((ABS(MAX(C67:C72)-B62)),(ABS(MIN(C67:C72)-B62)))),3)"
  objExcel.Range("D73").Formula = "=ROUND(ABS(MAX((ABS(MAX(D67:D72)-B63)),(ABS(MIN(D67:D72)-B63)))),3)"
  objExcel.Range("A73:E73").Font.Bold = True
  objExcel.Range("E74").Value = "Wiederholgenauigkeit"
  objExcel.Range("A74").Formula = "=ABS(MAX(A67:A72)-MIN(A67:A72))"
  objExcel.Range("B74").Formula = "=ABS(MAX(B67:B72)-MIN(B67:B72))"
  objExcel.Range("C74").Formula = "=ABS(MAX(C67:C72)-MIN(C67:C72))"
  objExcel.Range("D74").Formula = "=ABS(MAX(D67:D72)-MIN(D67:D72))"
  objExcel.Range("A74:E74").Font.Bold = True

  'Messen K2:
  objExcel.Range("E84").Value = "Max Abweichung"
  objExcel.Range("A84").Formula = "=MAX(ABS(MIN(A78:A83)),MAX(A78:A83))"
  objExcel.Range("B84").Formula = "=MAX(ABS(MIN(B78:B83)),MAX(B78:B83))"
  objExcel.Range("C84").Formula = "=MAX(ABS(MIN(C78:C83)),MAX(C78:C83))"
  objExcel.Range("D84").Formula = "=MAX(ABS(MIN(D78:D83)),MAX(D78:D83))"
  objExcel.Range("A84:E84").Font.Bold = True
  objExcel.Range("E85").Value = "Wiederholgenauigkeit"
  objExcel.Range("A85").Formula = "=ABS(MAX(A78:A83)-MIN(A78:A83))"
  objExcel.Range("B85").Formula = "=ABS(MAX(B78:B83)-MIN(B78:B83))"
  objExcel.Range("C85").Formula = "=ABS(MAX(C78:C83)-MIN(C78:C83))"
  objExcel.Range("D85").Formula = "=ABS(MAX(D78:D83)-MIN(D78:D83))"
  objExcel.Range("A85:E85").Font.Bold = True

  'Aus- und Einspannen K2:
  objExcel.Range("E95").Value = "Max Abweichung"
  objExcel.Range("A95").Formula = "=MAX(ABS(MIN(A89:A94)),MAX(A89:A94))"
  objExcel.Range("B95").Formula = "=MAX(ABS(MIN(B89:B94)),MAX(B89:B94))"
  objExcel.Range("A95:E95").Font.Bold = True
  objExcel.Range("E96").Value = "Wiederholgenauigkeit"
  objExcel.Range("A96").Formula = "=ABS(MAX(A89:A94)-MIN(A89:A94))"
  objExcel.Range("B96").Formula = "=ABS(MAX(B89:B94)-MIN(B89:B94))"
  objExcel.Range("A96:E96").Font.Bold = True
End If

objExcel.Columns("A:E").HorizontalAlignment = -4131 'links
objExcel.ScreenUpdating = True

'Konstanten: http://woonjas.linuxnerd.org/web/download.nsf/files/8C6C00FB633BCDC0C1256F39001D899E/$file/msoffice_constants.txt
