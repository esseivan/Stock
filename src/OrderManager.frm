VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderManager 
   Caption         =   "Order Manager"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2130
   OleObjectBlob   =   "OrderManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Run
End Sub

Private Sub Run()
    columnMan = -1
    Select Case (ListBox1.List(ListBox1.ListIndex))
        Case "Digikey"
        columnMan = s_art_ic_Digikey
        Case "Farnell"
        columnMan = s_art_ic_Farnell
        Case "Distrelec"
        columnMan = s_art_ic_Distrelec
        Case "Conrad"
        columnMan = s_art_ic_Conrad
        Case "Mouser"
        columnMan = s_art_ic_Mouser
        Case "Aliexpress"
        columnMan = s_art_ic_Aliexpress
        Case "Banggood"
        columnMan = s_art_ic_Banggood
        Case Else
        columnMan = s_art_ic_Other
    End Select
    
    If (columnMan <> -1) Then
        CheckValues (columnMan)
    End If
End Sub

Public Sub AddEntry(counter As Integer, columnMan As Integer)
    Dim ws2 As Worksheet
    Set ws2 = Worksheets("Articles")
    Dim ws3 As Worksheet
    Set ws3 = Worksheets("Commands")
    Dim newRange As Range
    RowCount = ws3.UsedRange.Rows.count - 1
    Set newRange = ws3.Range(ws3.Cells(RowCount, 1), ws3.Cells(RowCount, ws3.UsedRange.Columns.count))
    If Not (newRange.Cells(1, 2).Value = Empty) Then
        RowCount = RowCount + 1
        ' Create next row
        ws3.Rows(ws3.UsedRange.Rows.count).Insert shift:=xlShiftDowne
        Set newRange = ws3.Range(ws3.Cells(RowCount, 1), ws3.Cells(RowCount, ws3.UsedRange.Columns.count))
    End If
    Dim oldRange As Range
    Set oldRange = ws2.Range(ws2.Cells(counter, 1), ws2.Cells(counter, ws2.UsedRange.Columns.count))
    
    ' Copy data
    ret_n = oldRange.Cells(1, columnMan).Value
    If Not (ret_n = Empty) Then
        ' Check if already in list
        Dim searchRange As Range
        Set searchRange = ws3.Range(ws3.Cells(s_cmd_ir_start, s_cmd_ic_ret_n), ws3.Cells(ws3.UsedRange.Rows.count, s_cmd_ic_ret_n))
        
        Dim articles_ret_nbr As Range
        Set articles_ret_nbr = searchRange.Find( _
            what:=ret_n, _
            lookat:=xlWhole, _
            LookIn:=xlValues)
        
        If Not articles_ret_nbr Is Nothing Then
            ' Add to quantity
            Dim rowRange As Range
            Set rowRange = ws3.Range(ws3.Cells(articles_ret_nbr.Row, 1), ws3.Cells(articles_ret_nbr.Row, s_cmd_ic_lastRow))
            currentqty = rowRange.Cells(1, s_cmd_ic_qty).Value
            minqty = oldRange.Cells(1, s_art_ic_min).Value
            stockqty = oldRange.Cells(1, s_art_ic_stock).Value
            If (currentqty - minqty + stockqty) < 0 Then
                ' Missing qty, setting to (minqty - stockqty)
                Set newRange = ws3.Range(ws3.Cells(rowRange.Row, 1), ws3.Cells(rowRange.Row, ws3.UsedRange.Columns.count))
                newRange.Cells(1, s_cmd_ic_qty).Value = WorksheetFunction.RoundUp(minqty - stockqty, 0)
            End If
        Else
            ' Create next row
            ws3.Rows(ws3.UsedRange.Rows.count).Insert shift:=xlShiftDowne
            ' Create new entry
            newRange.Cells(1, s_cmd_ic_art_n).Value = oldRange.Cells(1, s_art_ic_art_n).Value
            newRange.Cells(1, s_cmd_ic_ret_n).Value = ret_n
            newRange.Cells(1, s_cmd_ic_man).Value = oldRange.Cells(1, s_art_ic_man).Value
            newRange.Cells(1, s_cmd_ic_ret).Value = ListBox1.List(ListBox1.ListIndex)
            newRange.Cells(1, s_cmd_ic_place).Value = oldRange.Cells(1, s_art_ic_place).Value
            newRange.Cells(1, s_cmd_ic_desc).Value = oldRange.Cells(1, s_art_ic_desc).Value
            newRange.Cells(1, s_cmd_ic_stock).Value = oldRange.Cells(1, s_art_ic_stock).Value
            newRange.Cells(1, s_cmd_ic_min).Value = oldRange.Cells(1, s_art_ic_min).Value
            newRange.Cells(1, s_cmd_ic_price).Value = oldRange.Cells(1, columnMan + 1).Value
        End If
    End If
End Sub


Public Sub CheckValues(columnMan As Integer)
    Dim ws2 As Worksheet
    Set ws2 = Worksheets("Articles")
    Dim ws3 As Worksheet
    Set ws3 = Worksheets("Commands")
    
    lastRow = ws2.UsedRange.Rows.count
    j = 0
    Dim iCntr As Integer
    For iCntr = 2 To lastRow
        ' Read next order column
        nextOrder = ws2.Cells(iCntr, s_art_ic_nextOrder).Value
        If (nextOrder = 1) Then
            AddEntry counter:=iCntr, columnMan:=columnMan
            j = j + 1
        End If
    Next
    MsgBox "Complete !" & vbNewLine & j & " entries processed"
End Sub


Private Sub UserForm_Initialize()
    ListBox1.AddItem ("Digikey")
    ListBox1.AddItem ("Farnell")
    ListBox1.AddItem ("Distrelec")
    ListBox1.AddItem ("Conrad")
    ListBox1.AddItem ("Mouser")
    ListBox1.AddItem ("Aliexpress")
    ListBox1.AddItem ("Banggood")
    ListBox1.SetFocus
End Sub
Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
    If KeyAscii = 13 Then
        Run
    End If
End Sub
Private Sub CommandButton1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
