Public Sub Test_Basic()
    Application.CutCopyMode = False

    ConvertCellDateFormat "APPT_DATE_TIME"

    Dim arrColumns As Range
    Set arrColumns = Union( _
        GetColumnByHeader("LANGUAGE"), _
        GetColumnByHeader("APPT_STATS"), _
        GetColumnByHeader("APPT_TYPE"), _
        GetColumnByHeader("APPT_DATE_TIME"), _
        GetColumnByHeader("RESOURCE"), _
        GetColumnByHeader("DURATION"), _
        GetColumnByHeader("LOCATION") _
    )
    arrColumns.Select
    Selection.Copy
    
    Dim targetWorksheet As Worksheet
    Set targetWorksheet = CreateSheet("Rpt1")
    targetWorksheet.Paste
    targetWorksheet.Activate
    
    Dim filteredRows As Range
    Set filteredRows = GetRowsByKeyValuePair("APPT_STATS", "CONFIRMED", Sheets("Rpt1"))
    'MsgBox filteredRows.Rows.Count
    'Set filteredRows = FilterRangeByKeyValuePair(filteredRows, "LANGUAGE", "TAMIL")
    Set filteredRows = FilterRangeByKeyValuePair(filteredRows, "LANGUAGE", "PERSIAN")
    Set filteredRows = FilterRangeByKeyValuePair(filteredRows, "APPT_TYPE", "HCIS Phone  CPT - iPM Patient")
    
    'Selection.Clear
    
    filteredRows.Select
    Selection.Copy
    
    Set targetWorksheet = CreateSheet()
    targetWorksheet.Paste
    targetWorksheet.Activate
    
    Set arrColumns = Union( _
        GetColumnByHeader("LANGUAGE"), _
        GetColumnByHeader("APPT_STATS"), _
        GetColumnByHeader("APPT_TYPE"), _
        GetColumnByHeader("APPT_DATE_TIME"), _
        GetColumnByHeader("RESOURCE"), _
        GetColumnByHeader("DURATION"), _
        GetColumnByHeader("LOCATION") _
    )
    
    arrColumns.EntireColumn.AutoFit
    
    SortColumn "APPT_DATE_TIME"
    
    Application.CutCopyMode = True
End Sub

Public Sub Info()
    MsgBox Sheets(1).UsedRange.Columns(1).Rows.Count
End Sub

Public Function GetColumnByHeader(title As String, Optional ws As Worksheet, Optional includeHeader As Boolean = True) As Range
    If ws Is Nothing Then
        Set ws = Application.ActiveSheet
    End If

    For i = 1 To ws.UsedRange.Columns.Count
        If ws.Cells(1, i).Value = title Then
            If includeHeader = True Then
                Set GetColumnByHeader = ws.Range(ws.Cells(1, i), ws.Cells(ws.UsedRange.Columns(i).Rows.Count, i))
            Else
                Set GetColumnByHeader = ws.Range(ws.Cells(2, i), ws.Cells(ws.UsedRange.Columns(i).Rows.Count, i))
            End If
            Exit Function
        End If
    Next
End Function

Public Function CreateSheet(Optional name As String, Optional position As Integer = -1) As Worksheet
    Dim newWorksheet As Worksheet
    If position <> -1 Then
        Set newWorksheet = Sheets.Add(Sheets(position))
    Else
        Set newWorksheet = Sheets.Add(, Sheets(Sheets.Count))
    End If
    
    If name <> "" Then
        newWorksheet.name = name
    End If
    
    Set CreateSheet = newWorksheet
End Function

Public Function GetRowsByKeyValuePair(columnName As String, rowValue As String, Optional ws As Worksheet) As Range
    Dim matchedRange As Range
    
    If ws Is Nothing Then
        Set ws = Application.ActiveSheet
    End If
    
    Set matchedRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.UsedRange.Columns.Count))
    
    For c = 1 To ws.UsedRange.Columns.Count
        If ws.Cells(1, c).Value = columnName Then
            For r = 2 To ws.UsedRange.Rows.Count
                If ws.Cells(r, c).Value = rowValue Then
                    Set matchedRange = Union(matchedRange, Range(ws.Cells(r, 1), ws.Cells(r, ws.UsedRange.Columns.Count)))
                End If
            Next
        End If
    Next
    
    'MsgBox matchedRange.Rows.Count
    
    Set GetRowsByKeyValuePair = matchedRange
End Function

Public Function FilterRangeByKeyValuePair(rng As Range, columnName As String, rowValue As String) As Range
    ' THIS SOME HOW BREAKS EVERYTHING
    
    Dim matchedRange As Range
    
    'Set matchedRange = rng.Cells(1, 1).EntireRow
    Set matchedRange = rng.Range(rng.Cells(1, 1), rng.Cells(1, rng.Columns.Count))
    
    For c = 1 To rng.Columns.Count
        If rng.Cells(1, c).Value = columnName Then
            For a = 1 To rng.Areas.Count
                For r = 1 To rng.Areas(a).Rows.Count
                    If rng.Areas(a).Cells(r, c).Value = rowValue Then
                        Set matchedRange = Union(matchedRange, rng.Range(rng.Areas(a).Cells(r, 1), rng.Areas(a).Cells(r, rng.Columns.Count)))
                    End If
                Next
            Next
        End If
    Next
    
    Set FilterRangeByKeyValuePair = matchedRange
End Function

Public Function ConvertCellDateFormat(columnName As String)
    Dim selectedRange As Range
    Set selectedRange = GetColumnByHeader(columnName)
    selectedRange.NumberFormat = "@"
    For Each cell In selectedRange.Cells
        If cell.Row > 1 Then
            If cell.Value <> "" And IsNumeric(cell.Value) Then
                Dim parsedDate As String
                parsedDate = Mid(cell.Value, 1, 4) & "-" & Mid(cell.Value, 5, 2) & "-" & Mid(cell.Value, 7, 2) & "  " & Mid(cell.Value, 9, 2) & ":" & Mid(cell.Value, 11, 2)
                cell.Value = parsedDate
            Else
                Exit Function
            End If
        End If
    Next cell
End Function

Public Function SortColumn(columnName As String, Optional ws As Worksheet)
    If ws Is Nothing Then
        Set ws = Application.ActiveSheet
    End If
    
    Dim column As Range
    Set column = GetColumnByHeader(columnName, ws, False)
    
    If column.Rows.Count > 1 Then
        ws.Sort.SortFields.Clear
        ws.Sort.SortFields.Add Key:=column, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ws.Sort
            .SetRange column.EntireRow
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

End Function