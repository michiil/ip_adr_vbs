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