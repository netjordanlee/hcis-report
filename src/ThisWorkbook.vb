Option Explicit

Dim cControl As CommandBarButton

Private Sub Workbook_Open()

    Workbook_AddinInstall

End Sub

Private Sub Workbook_AddinInstall()
On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Reformat Cerner Date").Delete
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
        With cControl
            .Caption = "Reformat Cerner Date"
            .Style = msoButtonCaption
            .OnAction = "CernerTool.ParseCernerDateFormat"
        End With
    Application.CommandBars("Worksheet Menu Bar").Controls("Report Tool").Delete
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
        With cControl
            .Caption = "Report Tool"
            .Style = msoButtonCaption
            .OnAction = "ReportTool.Test_Basic"
        End With
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinUninstall()
On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Reformat Cerner Date").Delete
    Application.CommandBars("Worksheet Menu Bar").Controls("Report Tool").Delete
    On Error GoTo 0
End Sub