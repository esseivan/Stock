VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StockManager 
   Caption         =   "Remove product from stock"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   OleObjectBlob   =   "StockManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name : Stock manager
' Author : Esseiva Nicolas
' Last modification : 05.02.19
' Description : Allow the management of the stock
'  It allows to remove, add and set components
'  Components can be searched through :
'   Man. art. number
'   Ret. art. number
'   Description
'   Place

Option Compare Text
Option Explicit
Dim result As String
Dim lastIndex As Integer
' Update infos
Private Sub btnUpdate_Click()
    lblStatus.Caption = "Working..."
    lastIndex = ListBox1.ListIndex
    result = StockModule.EditArticleInfos( _
        ListBox1.List(ListBox1.ListIndex), _
        ListBox2.ListIndex, _
        mLbl_manart.text, _
        mLbl_man.text, _
        mLbl_place.text, _
        mLbl_desc.text, _
        mLbl_qty.text, _
        mLbl_auto.Value)
    ' Write in history
    WriteHistory_modify mLbl_manart.text, "Article infos updated"
    ' Update filter
    FillArticleList txtboxSearch.Value
    ListBox1.ListIndex = lastIndex
    lblStatus.Caption = "Complete"
End Sub

' Initialize
Private Sub UserForm_Initialize()
    ListBox2.AddItem "Art number"
    ListBox2.AddItem "Retailer. Art. number"
    ListBox2.AddItem "Description"
    ListBox2.AddItem "Place"
    ListBox2.Selected(0) = True
    
    FillArticleList ""
    ListBox1.Selected(0) = True
    txtboxSearch.SetFocus
End Sub

' Events
' Remove
Private Sub btnDel_Click()
    ' Remove from stock
    result = StockModule.EditArticle(1, txtboxQty.Value, ListBox1.List(ListBox1.ListIndex), ListBox2.ListIndex)
    ' Check result
    If (IsNumeric(result)) Then
        mLbl_stock.Value = result
        txtboxQty.text = ""
    Else
        MsgBox "Invalid value"
    End If
End Sub
' Add
Private Sub btnAdd_Click()
    ' Add to stock
    result = StockModule.EditArticle(2, txtboxQty.Value, ListBox1.List(ListBox1.ListIndex), ListBox2.ListIndex)
    ' Check result
    If (IsNumeric(result)) Then
        mLbl_stock.Value = result
        txtboxQty.text = ""
    Else
        MsgBox "Invalid value"
    End If
End Sub
' Set
Private Sub btnSet_Click()
    ' Set to stock
    result = StockModule.EditArticle(3, txtboxQty.Value, ListBox1.List(ListBox1.ListIndex), ListBox2.ListIndex)
    ' Check result
    If (IsNumeric(result)) Then
        mLbl_stock.Value = result
        txtboxQty.text = ""
    Else
        MsgBox "Invalid value"
    End If
End Sub

' Search box change event
Private Sub txtboxSearch_Change()
    ' Update filter
    FillArticleList txtboxSearch.Value
End Sub

' Method to run on selected product change
' Listbox product change event
Sub ListBox1_Change()
    If (ListBox1.ListIndex >= 0) Then
        lblStatus.Caption = ""
        Dim listbox_index As Integer
        listbox_index = ListBox2.ListIndex
        
        ' Articles sheet
        Dim wsa As Worksheet
        Set wsa = Sheets("Articles")
        
        ' Search and repalce range
        Dim searchRange, _
            replaceRange As Range
        
        
        Dim col_name As String
        If (listbox_index = 0) Then
            ' Search by man. n.
            col_name = StockModule.ColumnIndexToString(StockModule.s_art_ic_art_n)
            Set searchRange = wsa.Range(col_name & "2:" & col_name & wsa.UsedRange.Rows.count)
        Else
            If (listbox_index = 1) Then
                ' Search by art. n.
                col_name = StockModule.ColumnIndexToString(StockModule.s_art_ic_default_art)
                Set searchRange = wsa.Range(col_name & "2:" & col_name & wsa.UsedRange.Rows.count)
            Else
                If (listbox_index = 2) Then
                    ' Search by desc
                    col_name = StockModule.ColumnIndexToString(StockModule.s_art_ic_desc)
                    Set searchRange = wsa.Range(col_name & "2:" & col_name & wsa.UsedRange.Rows.count)
                Else
                    If (listbox_index = 3) Then
                        ' Search by place
                        col_name = StockModule.ColumnIndexToString(StockModule.s_art_ic_place)
                        Set searchRange = wsa.Range(col_name & "2:" & col_name & wsa.UsedRange.Rows.count)
                    End If
                End If
            End If
        End If
        
        ' Stock range
        col_name = StockModule.ColumnIndexToString(StockModule.s_art_ic_stock)
        Set replaceRange = wsa.Range(col_name & "2:" & col_name & wsa.UsedRange.Rows.count)
        
        ' Find match in range
        Dim articles_art_nbr As Range
        Set articles_art_nbr = searchRange.Find( _
            what:=ListBox1.List(ListBox1.ListIndex), _
            lookat:=xlWhole, _
            LookIn:=xlValues)
        
        ' If found continue
        If Not articles_art_nbr Is Nothing Then
            ' Row of the match
            Dim Article_Row As Integer
            Article_Row = articles_art_nbr.Row
            
            ' Display informations
            ' Man. art number
            mLbl_manart.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_art_n).Value
            ' Ret. art number
            mLbl_retart.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_default_art).Value
            ' Manufacturer
            mLbl_man.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_man).Value
            'Place
            mLbl_place.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_place).Value
            ' Description
            'mLbl_desc.Value = "Description : " & wsa.Cells(Article_Row, StockModule.s_art_ic_desc).Value
            mLbl_desc.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_desc).Value
            ' Stock
            mLbl_stock.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_stock).Value
            ' min qty
            mLbl_qty.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_min).Value
            ' auo
            mLbl_auto.Value = wsa.Cells(Article_Row, StockModule.s_art_ic_auto).Value
            
            ' Go to the match
            Application.Goto Worksheets("Articles").Rows(Article_Row), True
        End If
    End If
End Sub

' Listbox search change event
Private Sub ListBox2_Change()
    ' Change type of filter
    FillArticleList txtboxSearch.Value
End Sub

' Fill article list (with filter)
Private Sub FillArticleList(filter)
    ' Clear list
    ListBox1.Clear
    
    ' Articles sheet
    Dim wsa As Worksheet
    Set wsa = Sheets("Articles")
    ' Used range
    Dim articles_range As Range
    Set articles_range = wsa.UsedRange
    
    Dim column_list As Integer
    If (ListBox2.Selected(0)) Then
        ' art. n.
        column_list = StockModule.s_art_ic_art_n
    Else
        If (ListBox2.Selected(1)) Then
            ' ret. n.
            column_list = StockModule.s_art_ic_default_art
        Else
            If (ListBox2.Selected(2)) Then
                ' Desc
                column_list = StockModule.s_art_ic_desc
            Else
                If (ListBox2.Selected(3)) Then
                    ' Place
                    column_list = StockModule.s_art_ic_place
                End If
            End If
        End If
    End If
    
    ' Generate text and add to list
    Dim i As Integer
    For i = 2 To articles_range.Rows.count
        ' Get text
        Dim string_t As String
        string_t = articles_range.Cells(i, column_list).Value
        
        ' Check for filter
        If (string_t Like "*" & filter & "*") Then
            ' If filter valid, add to list
            ListBox1.AddItem pvargitem:=string_t
        End If
    Next i
    
    SortListBox ListBox1
    
    ' Select first entry
    If (ListBox1.ListCount > 0) Then
        ListBox1.Selected(0) = True
    End If
End Sub

' Sort listbox alphabetically
Sub SortListBox(lst)
    Dim arrItems As Variant
    Dim arrTemp As Variant
    Dim intOuter As Long
    Dim intInner As Long
 
    arrItems = lst.List
 
    For intOuter = LBound(arrItems, 1) To UBound(arrItems, 1)
        For intInner = intOuter + 1 To UBound(arrItems, 1)
            If arrItems(intOuter, 0) > arrItems(intInner, 0) Then
                arrTemp = arrItems(intOuter, 0)
                arrItems(intOuter, 0) = arrItems(intInner, 0)
                arrItems(intInner, 0) = arrTemp
            End If
        Next intInner
    Next intOuter
 
    'Clear the listbox
    lst.Clear
 
    'Add the sorted array back to the listbox
    For intOuter = LBound(arrItems, 1) To UBound(arrItems, 1)
        lst.AddItem arrItems(intOuter, 0)
    Next intOuter
End Sub

' KeyPress events
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub ListBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_manart_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_retart_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_man_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_ret_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_desc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_place_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub mLbl_price_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnUpdate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txtboxQty_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnDel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnAdd_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnSet_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txtboxSearch_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub



