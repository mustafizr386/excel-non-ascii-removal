VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveLeadingNumbersMenu 
   Caption         =   "RemoveLeadingNumbers"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   OleObjectBlob   =   "RemoveLeadingNumbersMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveLeadingNumbersMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub executeButton_Click()
    Dim currentCell As String
    Dim startPatt As String
    
    Dim startRow As Long
    Dim endRow As Long
    Dim inputColumn As Long
    
    startRow = Val(startBox.Text)
    endRow = Val(endBox.Text)
    inputColumn = Val(inputBox.Text)
    
    For counter = startRow To endRow
        currentCell = Cells(counter, inputColumn).Value
        If IsNumeric(Left(currentCell, 1)) Then
            startPatt = Trim(Left(WorksheetFunction.Substitute(currentCell, " ", String(Len(currentCell), " ")), Len(currentCell)))
            Cells(counter, inputColumn).Value = Right(currentCell, Len(currentCell) - Len(startPatt) - 1)
        ElseIf Left(currentCell, 1) = " " Then
            Cells(counter, inputColumn).Value = Right(currentCell, Len(currentCell) - 1)
        ElseIf Right(currentCell, 1) = " " Then
            Cells(counter, inputColumn).Value = Left(currentCell, Len(currentCell) - 1)
        End If
        
    Next counter
    Unload Me
End Sub
