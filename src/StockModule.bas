Attribute VB_Name = "StockModule"
' Aricles sheet
Public Const s_art_ir_start = 2
Public Const s_art_ic_art_n = 1
Public Const s_art_ic_man = 2
Public Const s_art_ic_place = 3
Public Const s_art_ic_desc = 4
Public Const s_art_ic_stock = 5
Public Const s_art_ic_min = 6
Public Const s_art_ic_auto = 7
Public Const s_art_ic_default_art = 8
Public Const s_art_ic_default_price = 9
Public Const s_art_ic_Digikey = 10
Public Const s_art_ic_Farnell = 12
Public Const s_art_ic_Distrelec = 14
Public Const s_art_ic_Conrad = 16
Public Const s_art_ic_Mouser = 18
Public Const s_art_ic_Aliexpress = 20
Public Const s_art_ic_Banggood = 22
Public Const s_art_ic_Other = 24
Public Const s_art_ic_nextOrder = 26
Public Const s_art_ic_lastColumn = 26

' Product sheet
Public Const s_product_ir_start = 2
Public Const s_product_ic_art_n = 1
Public Const s_product_ic_desc = 2
Public Const s_product_ic_ref = 3
Public Const s_product_ic_qty = 4
Public Const s_product_ic_stock = 5
Public Const s_product_ic_available = 6
Public Const s_product_ic_t_stock = 7
Public Const s_product_ic_t_available = 8
Public Const s_product_ic_multiplier = 11
Public Const s_product_ic_lastColumn = 11

' Product config sheet
Public Const s_pc_ir_sheets = 1
Public Const s_pc_ir_start = 7

' Order adder config
Public Const s_oa_ir_i_start = 5
Public Const s_oa_ic_i_qty = 1
Public Const s_oa_ic_i_art_n = 2
Public Const s_oa_ic_i_man_n = 3
Public Const s_oa_ic_i_man = 4
Public Const s_oa_ic_i_ret = 5
Public Const s_oa_ic_i_desc = 6
Public Const s_oa_ic_i_price = 7
Public Const s_oa_ic_i_srow = 8

' Stock history
Public Const s_sh_ic_date = 1
Public Const s_sh_ic_article = 2
Public Const s_sh_ic_modified = 3
Public Const s_sh_ic_before = 4
Public Const s_sh_ic_after = 5
Public Const s_sh_ic_info = 6

' Commands
Public Const s_cmd_ir_start = 2
Public Const s_cmd_ic_art_n = 1
Public Const s_cmd_ic_ret_n = 2
Public Const s_cmd_ic_man = 3
Public Const s_cmd_ic_ret = 4
Public Const s_cmd_ic_place = 5
Public Const s_cmd_ic_desc = 6
Public Const s_cmd_ic_stock = 7
Public Const s_cmd_ic_min = 8
Public Const s_cmd_ic_price = 9
Public Const s_cmd_lastColumnAuto = 9
Public Const s_cmd_ic_qty = 10
Public Const s_cmd_ic_total = 11
Public Const s_cmd_ic_lastRow = 11

' Articles Command List


' Convert a column index to the column name
Function ColumnIndexToString(index) As String
    ' Get the range name (e.g. C1)
    x = Cells(1, index).Address(False, False)
    ' Remove the last character (in every case : "1", because we select row 1)
    ColumnIndexToString = Mid(x, 1, Len(x) - 1)
End Function

' Process an order operation
Sub ProcessOrder(error, operation, info, Optional force)
    ' Get index
    Set ws_oa = Worksheets("OrderAdder")
    Set ws_oaw = Worksheets("OrderAdder_work")
        
    ' Get config
    index_qty = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_qty) & s_oa_ir_i_start).Value
    index_art_n = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_art_n) & s_oa_ir_i_start).Value
    index_man_n = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_man_n) & s_oa_ir_i_start).Value
    index_man = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_man) & s_oa_ir_i_start).Value
    index_ret = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_ret) & s_oa_ir_i_start).Value
    index_desc = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_desc) & s_oa_ir_i_start).Value
    index_up = ws_oa.Range(ColumnIndexToString(s_oa_ic_i_price) & s_oa_ir_i_start).Value
    index_start = CInt(ws_oa.Range(ColumnIndexToString(s_oa_ic_i_srow) & s_oa_ir_i_start).Value)
    
    If (IsMissing(force)) Then
        force = False
    End If
    
    ' Check if numeric
    state_qty = IsNumeric(index_qty) And Not IsEmpty(index_qty)
    If Not state_qty Then
        ' If not numeric, save string in value
        value_qty = index_qty
    Else
        ' If numeric, convert and save in index
        index_qty = CInt(index_qty)
    End If
    
    ' Repeat for others parameters
    state_art_n = IsNumeric(index_art_n) And Not IsEmpty(index_art_n)
    If Not state_art_n Then
        value_art_n = index_art_n
    Else
        index_art_n = CInt(index_art_n)
    End If
    
    ' Repeat for others parameters
    state_man_n = IsNumeric(index_man_n) And Not IsEmpty(index_man_n)
    If Not state_art_n Then
        value_man_n = index_man_n
    Else
        index_man_n = CInt(index_man_n)
    End If
    
    ' Repeat for others parameters
    state_man = IsNumeric(index_man) And Not IsEmpty(index_man)
    If Not state_man Then
        value_man = index_man
    Else
        index_man = CInt(index_man)
    End If
    
    ' Repeat for others parameters
    state_ret = IsNumeric(index_ret) And Not IsEmpty(index_ret)
    If Not state_ret Then
        value_ret = index_ret
    Else
        index_ret = CInt(index_ret)
    End If
    
    ' Repeat for others parameters
    state_desc = IsNumeric(index_desc) And Not IsEmpty(index_desc)
    If Not state_desc Then
        value_desc = index_desc
    Else
        index_desc = CInt(index_desc)
    End If
    
    ' Repeat for others parameters
    state_up = IsNumeric(index_up) And Not IsEmpty(index_up)
    If Not state_up Then
        value_up = index_up
    Else
        index_up = CInt(index_up)
    End If
    
    ' Process order
    ' Get row count
    RowCount = ws_oaw.UsedRange.Rows.count
    counter = index_start
    ' Process each line
    j = 0
    While (counter <= RowCount)
        WriteLog "Counter : " & counter
        
        ' If is numeric (is a row index), get the value
        If state_qty Then
            value_qty = ws_oaw.Cells(counter, index_qty).Value
        End If
        
        ' Repeat for others parameters
        If state_art_n Then
            value_art_n = ws_oaw.Cells(counter, index_art_n).Value
        End If
        
        ' Repeat for others parameters
        If state_man_n Then
            value_man_n = ws_oaw.Cells(counter, index_man_n).Value
        End If
        
        ' Repeat for others parameters
        If state_man Then
            value_man = ws_oaw.Cells(counter, index_man).Value
        End If
        
        ' Repeat for others parameters
        If state_ret Then
            value_ret = ws_oaw.Cells(counter, index_ret).Value
        End If
        
        If state_desc Then
            value_desc = ws_oaw.Cells(counter, index_desc).Value
        End If
        
        ' Repeat for others parameters
        If state_up Then
            value_up = ws_oaw.Cells(counter, index_up).Value
        End If
        
        If (force = True) Then
            WriteLog "Article creation forced : " & value_man_n & ". Creating entry"
            StockModule.AddArticle _
                art_n:=CStr(value_art_n), _
                ret_n:=CStr(value_ret_n), _
                man:=value_man, _
                ret:=value_ret, _
                desc:=value_desc, _
                qty:=value_qty, _
                up:=value_up, _
                place:="None", _
                showError:=error
                j = j + 1
        Else
        ' Edit article
        If (EditArticle(operation, value_qty, value_man_n, 0, info) = "NAN") Then
            ' Article not existing
            WriteLog "Article not found : " & value_man_n & ". Creating entry"
            StockModule.AddArticle _
                art_n:=CStr(value_man_n), _
                ret_n:=CStr(value_art_n), _
                man:=value_man, _
                ret:=value_ret, _
                desc:=value_desc, _
                qty:=value_qty, _
                up:=value_up, _
                place:="None", _
                showError:=error
                j = j + 1
        End If
        End If
        
        ' increment counter
        counter = counter + 1
    Wend
    ' log
    WriteLog "Complete !" & vbNewLine & j & " articles created" & vbNewLine & (RowCount - j - index_start + 1) & " articles edited"
    MsgBox "Complete !" & vbNewLine & j & " articles created" & vbNewLine & (RowCount - j - index_start + 1) & " articles edited"
End Sub

' Convert string to double
Function ConvertToDouble(text) As Double
    ' Set regex to determine if convert is possible
    Set objRegExp = CreateObject("vbscript.regexp")
    With objRegExp
         .IgnoreCase = True
         .MultiLine = False
         .Pattern = "^[-]?[0-9]{1,}(\,|\.)?[0-9]{0,}?([E][-]?[0-9]{0,})?$"
         .Global = True
     End With
    
    ' If regex match
    If (objRegExp.test(text) = True) Then
        ' Replace , by .
        strVal = Trim(Replace(text, ",", "."))
        ConvertToDouble = Round(CDbl(strVal), 3)
    Else
    Err.Raise 1, "StockManager::ConvertToDouble()", "Quantity must be a number"
    End If
End Function

' Add article in Articles sheet
Function AddArticle(art_n, ret_n, man, ret, desc, qty, up, place, showError)
    ' Check if parameters are valid
    If (IsEmpty(art_n) Or IsNull(art_n) Or art_n = "") Then
        ' Invalid article, no man. number
        If (showError = True) Then
            MsgBox "One invalid article found : " & vbCrLf & "art_n = " & art_n & vbCrLf & _
                "ret_n = " & ret_n & vbCrLf & _
                "man = " & man & vbCrLf & _
                "ret = " & ret & vbCrLf & _
                "desc_n = " & desc & vbCrLf & _
                "qty = " & qty & vbCrLf & _
                "up = " & up & vbCrLf & _
                "place = " & place
        End If
    
        WriteLog "Invalid article found : \nart_n = " & art_n & _
            "\nret_n = " & ret_n & _
            "\nman = " & man & _
            "\nret = " & ret & _
            "\ndesc_n = " & desc & _
            "\nqty = " & qty & _
            "\nup = " & up & _
            "\npalce" = " & place"
    Else
        ' Get articles sheet
        Set ws = Sheets("Articles")
        ' Get range on Articles
        lastRow = ws.UsedRange.Rows.count
        Set lastArticle = ws.Range("A" & (lastRow + 1) & ":" & ColumnIndexToString(s_art_ic_lastColumn) & (lastRow + 1))
        
        ' Check if range is correct, cell must be empty
        If Not IsEmpty(lastArticle.Cells(1, 1).Value) Then
            MsgBox "Unknown error"
        Else
            columnMan = -1
            Select Case (ret)
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
                ' Copy values
                ' Art. number
                lastArticle.Cells(1, s_art_ic_art_n).Value = art_n
                ' Man
                lastArticle.Cells(1, s_art_ic_man).Value = man
                ' Place
                lastArticle.Cells(1, s_art_ic_place).Value = place
                ' Description
                lastArticle.Cells(1, s_art_ic_desc).Value = desc
                ' Quantity
                If IsNumeric(qty) And Not IsEmpty(qty) Then
                    lastArticle.Cells(1, s_art_ic_stock).Value = ConvertToDouble(qty)
                End If
                ' Min stock
                If IsNumeric(minStock) And Not IsEmpty(minStock) Then
                    lastArticle.Cells(1, s_art_ic_stock).Value = ConvertToDouble(minStock)
                End If
                ' Auto order
                If IsNumeric(auto) And Not IsEmpty(auto) Then
                    lastArticle.Cells(1, s_art_ic_stock).Value = ConvertToDouble(auto)
                End If
                
                ' Ret. number
                lastArticle.Cells(1, columnMan).Value = ret_n
                ' Unit price
                If IsNumeric(qty) And Not IsEmpty(qty) Then
                    lastArticle.Cells(1, columnMan + 1).Value = ConvertToDouble(up)
                End If
                
                ' Edit history
                WriteHistory art_n, qty, 0, qty, "Article added"
            Else
            If showError Then
                MsgBox "Invalid retailer"
            End If
            End If
        End If
    End If
End Function

' Edit article infos. Retailer art. number and retailer price has to be manually changed (for now)
Function EditArticleInfos(search, searchMode, art_n, man, place, desc, min, auto)
    Dim wsa As Worksheet
    Set wsa = Sheets("Articles")

    Dim searchRange As Range
    Set searchRange = GetSearchRange(searchMode)
    
    ' Find selected article in articles page
    Dim articles_found As Range
    Set articles_found = searchRange.Find( _
        what:=search, _
        lookat:=xlWhole, _
        LookIn:=xlValues)
    
    If Not articles_found Is Nothing Then
        ' Edit informations
        Dim wsa_line As Range
        Set wsa_line = wsa.Range(wsa.Cells(articles_found.Row, 1), wsa.Cells(articles_found.Row, s_art_ic_lastColumn))
        
        ' Edits
        ' art n
        wsa_line.Cells(1, s_art_ic_art_n).Value = art_n
        ' man
        wsa_line.Cells(1, s_art_ic_man).Value = man
        ' place
        wsa_line.Cells(1, s_art_ic_place).Value = place
        ' desc
        wsa_line.Cells(1, s_art_ic_desc).Value = desc
        ' min qty
        wsa_line.Cells(1, s_art_ic_min).Value = min
        ' auto order
        If (auto = True) Then
            auto = 1
        Else
            auto = 0
        End If
        wsa_line.Cells(1, s_art_ic_auto).Value = auto
    Else
        EditArticleInfos = "NAN"
    End If
    
End Function

' Get the search range from search mode
Function GetSearchRange(mode) As Range
    Set wsa = Sheets("Articles")
    Dim searchRange As Range
    
    If (mode = 0) Then
        ' art. number
        Set searchRange = wsa.Range(ColumnIndexToString(s_art_ic_art_n) & s_art_ir_start & ":" & ColumnIndexToString(s_art_ic_art_n) & wsa.UsedRange.Rows.count)
    Else
    If (mode = 1) Then
        ' Retailer number
        Set searchRange = wsa.Range(ColumnIndexToString(s_art_ic_Digikey) & s_art_ir_start & ":" & ColumnIndexToString(s_art_ic_Other) & wsa.UsedRange.Rows.count)
    Else
    If (mode = 2) Then
        ' Description
        Set searchRange = wsa.Range(ColumnIndexToString(s_art_ic_desc) & s_art_ir_start & ":" & ColumnIndexToString(s_art_ic_desc) & wsa.UsedRange.Rows.count)
    Else
    If (mode = 3) Then
        ' Place
        Set searchRange = wsa.Range(ColumnIndexToString(s_art_ic_place) & s_art_ir_start & ":" & ColumnIndexToString(s_art_ic_place) & wsa.UsedRange.Rows.count)
    End If
    End If
    End If
    End If
    
    Set GetSearchRange = searchRange
End Function

' Edit article
Function EditArticle(operation, quantity, text, searchMode, Optional info) As String
    Set wsa = Sheets("Articles")
    
    Set searchRange = GetSearchRange(searchMode)
    
    If (IsMissing(info)) Then
        txt_del = "Manual correction - Remove"
        txt_add = "Manual correction - Add"
        txt_set = "Manual correction - Set"
    Else
        txt_del = info
        txt_add = info
        txt_set = info
    End If
    
    ' Find selected article in articles page
    Set articles_found = searchRange.Find( _
        what:=CStr(text), _
        lookat:=xlWhole, _
        LookIn:=xlValues)
    
    ' If found continue
    If Not articles_found Is Nothing Then
        ' If quantity invalid, set to 0
        If (Not IsNumeric(quantity) Or quantity = "") Then
            quantity = 0
        End If
    
        ' Get art n
        art_nbr = wsa.Cells(articles_found.Row, s_art_ic_art_n).Value
        
        ' Search amount in stock
        Set articles_stock_range = wsa.Cells(articles_found.Row, s_art_ic_stock)
        
        ' convert to double
        articles_stock = ConvertToDouble(articles_stock_range.Value)
        articles_stock_before = articles_stock
        
        ' Edit the quantity
        If operation = 1 Then
            ' Substract
            articles_stock = Round(articles_stock - quantity, 3)
            WriteHistory art_nbr, -quantity, articles_stock_before, articles_stock, txt_del
        Else
            If operation = 2 Then
                ' Add
                articles_stock = Round(articles_stock + quantity, 3)
                WriteHistory art_nbr, quantity, articles_stock_before, articles_stock, txt_add
            Else
                If operation = 3 Then
                    ' Set
                    articles_stock = quantity
                    WriteHistory art_nbr, (quantity - articles_stock_before), articles_stock_before, articles_stock, txt_set
                Else
                    ' Unknown
                    MsgBox "ERROR ! Invalid operation !"
                    WriteLog "Unknown operation : " & operation
                End If
            End If
        End If
        
        ' Edit value in articles
        articles_stock_range.Value = articles_stock
        ' Return new stock
        EditArticle = articles_stock
    Else
        EditArticle = "NAN"
    End If
End Function

' Write history : Write entry in stock history sheet
Sub WriteHistory(art, count, before, after, info)
    ' Sheet
    Set ws = Sheets("StockHistory")
    ' Last row
    LS = ws.UsedRange.Rows.count + 1
    
    ' Edit cells
    ws.Cells(LS, s_sh_ic_date).Value = DateTime.Now
    ws.Cells(LS, s_sh_ic_article).Value = art
    ws.Cells(LS, s_sh_ic_modified).Value = count
    ws.Cells(LS, s_sh_ic_before).Value = before
    ws.Cells(LS, s_sh_ic_after).Value = after
    ws.Cells(LS, s_sh_ic_info).Value = info
End Sub

' Write history : Component modified
Sub WriteHistory_modify(art, info)
    ' Sheet
    Set ws = Sheets("StockHistory")
    ' Last row
    LS = ws.UsedRange.Rows.count + 1
    
    ' Edit cells
    ws.Cells(LS, s_sh_ic_date).Value = DateTime.Now
    ws.Cells(LS, s_sh_ic_article).Value = art
    ws.Cells(LS, s_sh_ic_modified).Value = "-"
    ws.Cells(LS, s_sh_ic_before).Value = "-"
    ws.Cells(LS, s_sh_ic_after).Value = "-"
    ws.Cells(LS, s_sh_ic_info).Value = info
End Sub

' Write log to console
Sub WriteLog(f)
    If Not f = "" Then
        Debug.Print f
    End If
End Sub
' Write details to console
Sub Report(f)
    If Not f Is Nothing Then
        Debug.Print f.Address & " | " & f
    Else
        Debug.Print "Nothing"
    End If
End Sub

