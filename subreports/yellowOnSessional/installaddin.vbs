On Error Resume Next
Dim oXL
Dim oAddin
Set oXL = CreateObject("Excel.Application")
oXL.Workbooks.Add
Set oAddin = oXL.AddIns.Add("T:\CSO\JL\tools\subreports\yellowOnSessional\ReportYellowOnSessional.xlam", True)
oAddin.Installed = True
oXL.Quit
Set oAddin = Nothing
Set oXL = Nothing

If Err.Number <> 0 Then
	MsgBox "Error: " & Err.Number & vbNewLine & "Srce: " & Err.Source & vbNewLine & "Desc: " &  Err.Description, vbOKOnly + vbCritical, "Encountered Error" 
Else
	MsgBox "Installed, please restart Excel" & vbNewLine & "The button can be found in the 'Add-Ins' option of the toolbar", vbOKOnly + vbInformation, "Successfully Installed" 
End If