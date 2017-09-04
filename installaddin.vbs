Dim Shell
Set Shell = WScript.CreateObject("WScript.Shell")

result = MsgBox ("This update requires Excel to close." & vbNewLine & vbNewLine & "ALL UNSAVED WORK WILL BE LOST!", vbOKCancel + vbExclamation, "ReportTool Update")

Select Case result
Case vbOK
    Shell.Run "taskkill /f /im ""EXCEL.EXE""", , true
Case vbCancel
    MsgBox "Update was cancelled", vbOKOnly + vbInformation, "ReportTool Update"
    WScript.Quit
End Select


Shell.Run "taskkill /f /im ""EXCEL.EXE""", , true

On Error Resume Next
Dim oXL
Dim oAddin
Set oXL = CreateObject("Excel.Application")
oXL.Workbooks.Add
Set oAddin = oXL.AddIns.Add("T:\CSO\JL\tools\ReportTool.xlam", True)
oAddin.Installed = True
oXL.Quit
Set oAddin = Nothing
Set oXL = Nothing

If Err.Number <> 0 Then
	MsgBox "Error: " & Err.Number & vbNewLine & "Srce: " & Err.Source & vbNewLine & "Desc: " &  Err.Description, vbOKOnly + vbCritical, "ReportTool Update" 
Else
	MsgBox "Installed, please restart Excel" & vbNewLine & "The button can be found in the 'Add-Ins' option of the toolbar", vbOKOnly + vbInformation, "ReportTool Update" 
End If