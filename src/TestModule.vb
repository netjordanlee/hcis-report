Public Function TestingMacro()
    
    Dim arrColumns As Range
    Set arrColumns = Union( _
        ReportTool.GetColumnByHeader("LANGUAGE"), _
        ReportTool.GetColumnByHeader("APPT_STATS"), _
        ReportTool.GetColumnByHeader("APPT_TYPE") _
    )
    arrColumns.Select
    Selection.Copy
    
    Dim targetWorksheet As Worksheet
    Set targetWorksheet = ReportTool.CreateSheet("MySpecialReport")
    
    targetWorksheet.Paste
    targetWorksheet.Activate

End Function

Public Function TestingBreak()
    
    ReportTool.ConvertCellDateFormat "APPT_DATE_TIME", ActiveWorkbook.Sheets(1)
    
    userInputDate = InputBox("Enter date in format yyyy-mm-dd", "Enter date to report on")
    Dim checkDate As Date
    checkDate = DateValue(userInputDate)
    
    Dim Interpreters() As Interpreter
    ReDim Interpreters(1)

    Dim wsMTLReport As Worksheet
    Set wsMTLReport = ReportTool.CreateSheet("MTL Report")

    ActiveWorkbook.Sheets(1).Activate ' Ensure MTL doesn't get focus
    
    Dim arrColumns As Range
    Set arrColumns = Union( _
        ReportTool.GetColumnByHeader("RESOURCE"), _
        ReportTool.GetColumnByHeader("APPT_DATE_TIME"), _
        ReportTool.GetColumnByHeader("APPT_STATS") _
    )
    
    ' Populate array of interpreters
    
    For c = 1 To arrColumns.Columns.Count
        If arrColumns.Cells(1, c).Value = "RESOURCE" Then
            For r = 1 To arrColumns.Rows.Count
                alreadyDetected = False
                'If UBound(Interpreters) = 1 Then
                '    Set Interpreters(UBound(Interpreters)) = New Interpreter
                '    Interpreters(UBound(Interpreters)).Resource = arrColumns.Cells(r, c).Value
                'End If
                For i = 1 To UBound(Interpreters)
                    If Interpreters(i) Is Nothing Then
                        Set Interpreters(UBound(Interpreters)) = New Interpreter
                        Interpreters(UBound(Interpreters)).Resource = arrColumns.Cells(r, c).Value
                        alreadyDetected = True
                    ElseIf Interpreters(i).Resource = arrColumns.Cells(r, c).Value Then
                        alreadyDetected = True
                    End If
                Next
                If alreadyDetected = False Then
                    ReDim Preserve Interpreters(UBound(Interpreters) + 1)
                    Set Interpreters(UBound(Interpreters)) = New Interpreter
                    Interpreters(UBound(Interpreters)).Resource = arrColumns.Cells(r, c).Value
                End If
            Next
        End If
    Next
    
    
    
    Set arrColumns = ReportTool.FilterRangeByKeyValuePair(arrColumns, "APPT_TYPE", "HCIS Face")
    
    For i = 1 To UBound(Interpreters)
        wsMTLReport.Cells(i, 1).Value = Interpreters(i).Resource
    Next
    
    For a = 1 To arrColumns.Areas.Count
        For r = 1 To arrColumns.Areas(a).Rows.Count
            For c = 1 To arrColumns.Columns.Count
                If arrColumns.Cells(1, c).Value = "RESOURCE" Then
                    
                End If
            Next
        Next
    Next
    
End Function
