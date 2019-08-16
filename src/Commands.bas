Attribute VB_Name = "Commands"
Public Function Clear()
    Dim wsC As Worksheet
    Set wsC = Worksheets("Commands")
    RowCount = wsC.UsedRange.Rows.count
    If (RowCount = 1) Then
        RowCount = 2
    End If
    
    ' Clear 2nd line (between header and total)
    wsC.Range(wsC.Cells(2, 1), wsC.Cells(3, s_cmd_lastColumnAuto)).ClearContents
    While (wsC.UsedRange.Rows.count > 4)
        wsC.Rows(4).Delete
    Wend
End Function

Public Sub ResetLines()
    Dim ws3 As Worksheet
    Set ws3 = Worksheets("Commands")
    Dim editRange As Range
    Set editRange = ws3.Range(ws3.Cells(2, 1), ws3.Cells(3, s_cmd_lastColumnAuto))
    
    editRange.Cells(1, s_cmd_ic_qty).Formula = "=ROUNDUP([@[Min Stock]]-[@[In stock]],0)"
    editRange.Cells(1, s_cmd_ic_total).Formula = "=[@Quantity]*[@[Unit price]]"
    editRange.Cells(2, s_cmd_ic_qty).Formula = "=ROUNDUP([@[Min Stock]]-[@[In stock]],0)"
    editRange.Cells(2, s_cmd_ic_total).Formula = "=[@Quantity]*[@[Unit price]]"
End Sub

Public Sub AddRow()
    Dim ws3 As Worksheet
    Set ws3 = Worksheets("Commands")
    ' Create next row
    ws3.Rows(ws3.UsedRange.Rows.count).Insert shift:=xlShiftDowne
End Sub
