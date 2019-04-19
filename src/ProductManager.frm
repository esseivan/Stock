VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProductManager 
   Caption         =   "Remove product from stock"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   OleObjectBlob   =   "ProductManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProductManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Name : Product manager
' Author : Esseiva Nicolas
' Last modification : 05.02.19
' Description :
'  Remove components from stock for a specified product if available
'  Undo possible (More like a add components rather than undo)

Dim count As Integer
Dim ws_pc As Worksheet

' Initialize
Private Sub UserForm_Initialize()
    ' Read data from config sheet
    Set ws_pc = Sheets("ProductConfig")
    
    ' Determine start row
    startrow = StockModule.s_pc_ir_start
    StockModule.WriteLog "Start row is : " & startrow
    
    ' Récupérer le tableau
    count = 0
    count2 = 0
    StockModule.WriteLog "Retrieving config..."
    ' If not empty
    Report ws_pc.Cells(startrow + count, 1)
    Do While (Not IsEmpty(ws_pc.Cells(startrow + count, 1).Value))
        ' If positive number
        If (ws_pc.Cells(startrow + count, 1).Value >= 0) Then
            ' Add to list
            StockModule.WriteLog "Index : " & ws_pc.Cells(startrow + count, 1).Value
            StockModule.WriteLog "Name : " & ws_pc.Cells(startrow + count, 2).Value
            StockModule.WriteLog "SheetName : " & ws_pc.Cells(startrow + count, 3).Value
            ListBox1.AddItem (ws_pc.Cells(startrow + count, 2).Value)
            ' Save as a unique row
            ws_pc.Cells(StockModule.s_pc_ir_sheets, count2 + 1).Value = ws_pc.Cells(startrow + count, 3).Value
            count2 = count2 + 1
        End If
        count = count + 1
        ' Log counter
        StockModule.WriteLog count
    Loop
    ' Log complete
    StockModule.WriteLog "Config complete !"
End Sub

' Remove button event
Private Sub btnRemove_Click()
    i = ListBox1.ListIndex
    
    If (i >= 0) Then
        RemoveFromStock (ws_pc.Cells(StockModule.s_pc_ir_sheets, i + 1).Value)
    End If
End Sub

' Undo button event
Private Sub btnUndo_Click()
    i = ListBox1.ListIndex
    
    If (i >= 0) Then
            UndoRemoveFromStock (ws_pc.Cells(StockModule.s_pc_ir_sheets, i + 1).Value)
    End If
End Sub

' Add components for the product from the stock
Private Sub UndoRemoveFromStock(i)
    Set ws_i = Sheets(i)
    
    Label1.BackColor = SystemColorConstants.vbButtonFace
    Label1.ForeColor = vbBlack
    Label1.Caption = "Added " & i & " to stock"
    
    RemoveProduct True, ws_i.UsedRange, i
End Sub

' Remove components for the product from the stock
Private Sub RemoveFromStock(i)
    Set ws = Sheets(i)
    
    Dim f As Range
    col_name_available = StockModule.ColumnIndexToString(StockModule.s_product_ic_available)
    Set f = ws.Range(col_name_available & "2:" & col_name_available & ws.UsedRange.Rows.count).Find( _
    what:="NO", _
    lookat:=xlWhole, _
    LookIn:=xlValues)
    
    ' If any "NO" found
    If Not f Is Nothing Then
        Label1.BackColor = vbRed
        Label1.ForeColor = vbWhite
        Label1.Caption = "Not all components are available Unable to remove"
    Else
        Label1.BackColor = vbGreen
        Label1.ForeColor = vbBlack
        Label1.Caption = "Removed " & i & " from stock"
        ' Else remove from stock
        RemoveProduct False, ws.UsedRange, i
    End If
End Sub

' Remove product from stock (1 time)
Private Sub RemoveProduct(undo, productRange, name)
    ' Get the articles sheet
    Set wsa = Sheets("Articles")
    ' Get the search range from articles sheet (man. art. number)
    col_name_art_n = StockModule.ColumnIndexToString(StockModule.s_art_ic_art_n)
    Set searchRange = wsa.Range(col_name_art_n & "2:" & col_name_art_n & wsa.UsedRange.Rows.count)
    ' Get the replace range from article sheet (stock)
    col_name_stock = StockModule.ColumnIndexToString(StockModule.s_art_ic_stock)
    Set replaceRange = wsa.Range(col_name_stock & "2:" & col_name_stock & wsa.UsedRange.Rows.count)
    ' Get the row count
    Max = productRange.Rows.count
    
    ' Process each line of the product
    For i = StockModule.s_product_ir_start To Max
        ' Find the art. number in the product sheet
        product_art_nbr = productRange.Cells(i, StockModule.s_product_ic_art_n)
        
        ' Find the total quantity in the product page
        product_quantity = productRange.Cells(i, StockModule.s_product_ic_qty).Value
        
        ' If quantity valid
        If (product_quantity >= 0) Then
            ' Find identical art. number in articles sheet
            Set articles_art_nbr = searchRange.Find( _
                what:=product_art_nbr, _
                lookat:=xlWhole, _
                LookIn:=xlValues)
            
            ' If found continue
            If Not articles_art_nbr Is Nothing Then
                ' Search amount in stock
                Set articles_stock_range = replaceRange((articles_art_nbr.Row - 1), 1)
                
                ' convert to double
                articles_stock = articles_stock_range.Value
                articles_stock_before = articles_stock
                
                ' Add or remove the quantity
                If undo Then
                    ' Add the quantity (undo: True)
                    articles_stock = articles_stock + (1 * product_quantity)
                    ' Log in history
                    WriteHistory product_art_nbr, (-1 * product_quantity), articles_stock_before, articles_stock, name
                Else
                    ' Remove the quantity (undo: False)
                    articles_stock = articles_stock - (1 * product_quantity)
                    ' Log in history
                    WriteHistory product_art_nbr, (1 * product_quantity), articles_stock_before, articles_stock, name
                End If
                
                ' Edit value in articles sheet
                articles_stock_range.Value = articles_stock
            End If
        End If
    Next i
End Sub

' ListBox change event
Private Sub ListBox1_Change()
    ' Check selected listbox index
    i = ListBox1.ListIndex
    
    If (i >= 0) Then
        ' Get the sheet of the selected product (from the productConfig sheet)
        Set wsa = Sheets(ws_pc.Cells(StockModule.s_pc_ir_sheets, i + 1).Value)
        
        ' Get the range of the available
        Dim f As Range
        col_name_available = StockModule.ColumnIndexToString(StockModule.s_product_ic_available)
        Set f = wsa.Range(col_name_available & "2:" & col_name_available & wsa.UsedRange.Rows.count).Find( _
        what:="NO", _
        lookat:=xlWhole, _
        LookIn:=xlValues)
        
        ' If any "NO" found
        If Not f Is Nothing Then
            Label1.BackColor = vbRed
            Label1.ForeColor = vbWhite
            Label1.Caption = "Not all components are available"
        Else
            Label1.BackColor = vbGreen
            Label1.ForeColor = vbBlack
            Label1.Caption = "Components available"
        End If
    End If
End Sub

' KeyPress events
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnRemove_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnUndo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub


