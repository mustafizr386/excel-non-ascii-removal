VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LengthExportMenu 
   Caption         =   "LengthExport"
   ClientHeight    =   2400
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "LengthExportMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LengthExportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub endTitle_Click()

End Sub

Private Sub executeButton_Click()
    Dim currentCell As String
    
    Dim startRow As Long
    Dim endRow As Long
    Dim maxLength As Integer
    Dim inputColumn As Long
    Dim r, g, b As Integer
    Dim internalCounter As Integer
    
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    inputColumn = Val(inputBox.Text)
    maxLength = Val(lengthBox.Text)
    
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveWorkbook.ActiveSheet
    
    Dim resultSheet As Worksheet
    Set resultSheet = ActiveWorkbook.Sheets.Add(After:= _
                    ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count))
    
    Dim rand As Integer
    rand = CInt(Int((100 * rnd())))
    
    resultSheet.Name = currentSheet.Name + "-LeXPORT"
    
    
    internalCounter = 2
    
    resultSheet.Cells(1, 1).Value = "Row Number"
    resultSheet.Cells(1, 2).Value = "Invalid Data"
    
    For counter = startRow To endRow
        currentCell = currentSheet.Cells(counter, inputColumn)
        If Len(currentCell) > maxLength Then
            resultSheet.Cells(internalCounter, 1).Value = counter
            resultSheet.Cells(internalCounter, 2).Value = currentCell
            internalCounter = internalCounter + 1
        End If
    Next counter
    Unload Me
End Sub
