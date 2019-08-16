VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderAdder 
   Caption         =   "Add order"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "OrderAdder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderAdder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name : Order adder
' Author : Esseiva Nicolas
' Last modification : 05.02.19
' Description :
'  Add an order to the stock
'  Add the quantity from existing articles
'  Create non existing articles

Dim FileName As Variant
Option Explicit

' Initialize
Private Sub UserForm_Initialize()
    lblStatus.Caption = ""
    Label1.Caption = "No file selected"
End Sub

' Load file
Private Sub loadFile()
    lblStatus.Caption = ""
    GetImportFileName
End Sub
' Load file 2
Sub GetImportFileName()
  Dim Finfo As String
  Dim FilterIndex As Long
  Dim Title As String
  
'  Set up list of file filters
  Finfo = "Excel Files (*.xl*),*.xl*," & _
        "Comma Separated Files (*.csv),*.csv," & _
        "All Files (*.*),*.*"
'  Display *.* by default
  FilterIndex = 1
'  Set the dialog box caption
 Title = "Select a File to Import"
'  Get the filename
  FileName = Application.GetOpenFilename(Finfo, _
    FilterIndex, Title)
'  Handle return info from dialog box
  If FileName = False Then
    StockModule.WriteLog "No file selected"
    Label1.Caption = "No file selected"
    FileName = ""
  Else
    StockModule.WriteLog "Selected file : " & FileName
    Label1.Caption = FileName
  End If
End Sub
' Open excel file
Private Sub Workbook_Open()
    Call ReadDataFromCloseFile
End Sub
' Open excel file 2
Sub ReadDataFromCloseFile()
    If (IsEmpty(FileName) Or FileName = "") Then
        MsgBox "No file selected !"
    Else
        On Error GoTo ErrHandler
        'Application.ScreenUpdating = False
            
        Dim BaseWorkbook As Workbook
        Set BaseWorkbook = ThisWorkbook
        BaseWorkbook.Worksheets("OrderAdder_work").UsedRange.Clear
            
        Dim src As Workbook
        
        ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
        Set src = Workbooks.Open(FileName, True, True)
        
        StockModule.WriteLog "Start"
        ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK.
        Dim sourceRange As Range
        Set sourceRange = src.Worksheets(1).UsedRange
        StockModule.WriteLog "Rows : " & sourceRange.Rows.count
        
        ' COPY DATA FROM SOURCE (CLOSE WORKGROUP) TO THE DESTINATION WORKBOOK.
        sourceRange.Copy
        BaseWorkbook.Worksheets("OrderAdder_work").Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        StockModule.WriteLog "complete"
        
        ' CLOSE THE SOURCE FILE.
        src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
        Set src = Nothing
        
        StockModule.WriteLog "Closed"
        Application.ScreenUpdating = True
        
ErrHandler:
        StockModule.WriteLog "Error"
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

' Events
' Load file
Private Sub btnChooseFile_Click()
    loadFile
    runReader
End Sub

' Clear sheet
Private Sub btnClearPage_Click()
    runCleaner
End Sub

Private Sub runCleaner()
    Dim ws As Worksheet
    Set ws = Worksheets("OrderAdder_work")
    Dim RowCount As Integer
    RowCount = ws.UsedRange.Rows.count
    ws.UsedRange.Delete
    lblStatus.Caption = "Clear complete ! " & RowCount & " row(s) deleted"
End Sub

' Processs order
Private Sub btnProcessOrder_Click()
    lblStatus.Caption = "Processing order..."
    StockModule.ProcessOrder CheckBox1.Value, 2, txtInfo.text, CheckBox2.Value
    lblStatus.Caption = "Process complete !"
End Sub

' Undo process order
Private Sub btnUndoProcessOrder_Click()
    lblStatus.Caption = "Undoing..."
    StockModule.ProcessOrder CheckBox1.Value, 1, txtInfo.text
    lblStatus.Caption = "Undo complete !"
End Sub

' Read file
Private Sub btnReadFile_Click()
    runReader
End Sub

Private Sub runReader()
    runCleaner
    Workbook_Open
    Worksheets("OrderAdder").Select
    lblStatus.Caption = "Read complete !"
End Sub

' KeyPress events
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnChooseFile_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnReadFile_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnProcessOrder_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnClearPage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txtInfo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub btnUndoProcessOrder_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub CheckBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub CheckBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Me
End Sub



