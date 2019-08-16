VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArticleAdder 
   Caption         =   "Add article"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ArticleAdder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArticleAdder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Name : Article adder
' Author : Esseiva Nicolas
' Last modification : 05.02.19
' Description :
'  Add an article to the list


' Todo : -

' OnLoad
Private Sub UserForm_Initialize()
    ' Set status label
    lblStatus.Caption = "Idle"
    ComboBox1.AddItem ("Digikey")
    ComboBox1.AddItem ("Farnell")
    ComboBox1.AddItem ("Distrelec")
    ComboBox1.AddItem ("Conrad")
    ComboBox1.AddItem ("Mouser")
    ComboBox1.AddItem ("Aliexpress")
    ComboBox1.AddItem ("Banggood")
    ComboBox1.ListIndex = 0
End Sub

' Add entry
Private Sub btnAdd_Click()
    ' Create new article
    If (ComboBox1.ListIndex <> -1) Then
        retailer = ComboBox1.List(ComboBox1.ListIndex)
    Else
        retailer = ComboBox1.text
    End If
    
    StockModule.AddArticle _
    art_n:=TextBox1.Value, _
    ret_n:=TextBox2.Value, _
    man:=TextBox3.Value, _
    ret:=retailer, _
    desc:=TextBox5.Value, _
    qty:=TextBox6.Value, _
    up:=TextBox7.Value, _
    place:=TextBox8.Value, _
    showError:=True
    
    ' Set status label
    lblStatus.Caption = "Complete"
    
    ' Focus first textbox
    TextBox1.SetFocus
    TextBox1.SelStart = 0
    TextBox1.SelLength = TextBox1.TextLength
End Sub

' Clear entries
Private Sub btnClear_Click()
    ' Set status label
    lblStatus.Caption = "Cleared"
    
    ' Clear all textboxes
    TextBox1.Value = Empty
    TextBox2.Value = Empty
    TextBox3.Value = Empty
    TextBox4.Value = Empty
    TextBox5.Value = Empty
    TextBox6.Value = Empty
    TextBox7.Value = Empty
    
    ' Focus first textbox
    TextBox1.SetFocus
End Sub

' KeyPress events
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lblStatus.Caption = ""
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnAdd_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnClear_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub




